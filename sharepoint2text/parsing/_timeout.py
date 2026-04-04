"""Utilities for enforcing extraction wall-clock time limits."""

from __future__ import annotations

import signal
import threading
import time
from contextlib import contextmanager
from math import isfinite
from typing import Any, BinaryIO, Callable, Generator, Iterator

from sharepoint2text.parsing.exceptions import ExtractionFailedError
from sharepoint2text.parsing.extractors.data_types import ExtractionInterface

DEFAULT_EXTRACTION_TIMEOUT_SECONDS = 60.0

ExtractorFunction = Callable[
    [BinaryIO, str | None], Generator[ExtractionInterface, Any, None]
]


def normalize_timeout_seconds(timeout_seconds: float | int) -> float:
    """Validate and normalize a timeout value.

    Args:
        timeout_seconds: Requested timeout in seconds. ``0`` disables the limit.

    Returns:
        The timeout as a ``float`` for downstream timer APIs.

    Raises:
        ValueError: If ``timeout_seconds`` is negative or not finite.
    """
    normalized = float(timeout_seconds)
    if normalized < 0:
        raise ValueError("timeout_seconds must be >= 0")
    if not isfinite(normalized):
        raise ValueError("timeout_seconds must be a finite number")
    return normalized


@contextmanager
def extraction_timeout(seconds: float, *, source: str) -> Iterator[None]:
    """Enforce a wall-clock timeout for extraction work.

    Args:
        seconds: Timeout in seconds. ``0`` disables the limit.
        source: Human-readable source label used in error messages.

    Raises:
        ExtractionFailedError: If timeout enforcement cannot run on the current
            platform or thread.
        TimeoutError: If the timeout expires while the wrapped work is running.
    """
    if seconds <= 0:
        yield
        return

    if not hasattr(signal, "setitimer") or not hasattr(signal, "SIGALRM"):
        raise ExtractionFailedError(
            f"Extraction timeout is not supported on this platform for {source}"
        )

    if threading.current_thread() is not threading.main_thread():
        raise ExtractionFailedError(
            f"Extraction timeout requires the main thread for {source}"
        )

    def _handle_timeout(_signum: int, _frame: object | None) -> None:
        raise TimeoutError(f"Extraction timed out after {seconds:g} seconds: {source}")

    previous_handler = signal.getsignal(signal.SIGALRM)
    start_time = time.monotonic()
    signal.signal(signal.SIGALRM, _handle_timeout)
    previous_timer = signal.setitimer(signal.ITIMER_REAL, seconds)

    try:
        yield
    finally:
        elapsed = time.monotonic() - start_time
        signal.setitimer(signal.ITIMER_REAL, 0.0)
        signal.signal(signal.SIGALRM, previous_handler)

        previous_delay, previous_interval = previous_timer
        remaining_delay = max(previous_delay - elapsed, 0.0)
        if remaining_delay > 0 or previous_interval > 0:
            signal.setitimer(
                signal.ITIMER_REAL,
                remaining_delay,
                previous_interval,
            )


def collect_extraction_results(
    extractor: ExtractorFunction,
    file_like: BinaryIO,
    path: str | None,
    *,
    timeout_seconds: float,
) -> list[ExtractionInterface]:
    """Consume an extractor generator while enforcing an optional timeout.

    Results are materialized before being yielded to callers so the timeout only
    measures extractor execution time, not downstream consumer processing time
    between yielded results.

    Args:
        extractor: Extractor callable to execute.
        file_like: Binary stream passed to the extractor.
        path: Source path shown in error messages and metadata.
        timeout_seconds: Timeout in seconds. ``0`` disables the limit.

    Returns:
        A list of extraction results emitted by the extractor.

    Raises:
        ExtractionFailedError: If the timeout expires.
    """
    source = path or "<in-memory>"
    try:
        with extraction_timeout(timeout_seconds, source=source):
            return list(extractor(file_like, path))
    except TimeoutError as exc:
        raise ExtractionFailedError(str(exc), cause=exc) from exc
