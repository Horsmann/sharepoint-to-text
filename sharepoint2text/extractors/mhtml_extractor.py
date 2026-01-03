"""
MHTML Web Archive Extractor
===========================

Extracts text content from MHTML (MIME HTML) web archive files.

File Format Background
----------------------
MHTML (MIME Encapsulation of Aggregate HTML Documents) is a web archive
format that bundles an HTML document with its resources (images, CSS, etc.)
into a single file using MIME multipart encoding. Common extensions:
    - .mhtml
    - .mht

MHTML files are structured as MIME messages:
    - Content-Type: multipart/related
    - First part: The main HTML document
    - Subsequent parts: Referenced resources (images, stylesheets, scripts)

The format is defined in RFC 2557 and commonly used by:
    - Microsoft Internet Explorer / Edge "Save as Web Archive"
    - Microsoft Word "Save as Single File Web Page"
    - Opera browser
    - SharePoint exports

Structure
---------
MHTML files contain:
    1. MIME headers (From, Subject, Date, Content-Type, etc.)
    2. Boundary-separated parts:
       - Main HTML content (text/html)
       - Embedded resources (image/*, text/css, etc.)

This extractor focuses on the HTML content, extracting it from the MIME
structure and processing it with the standard HTML extractor.

Dependencies
------------
Python Standard Library only:
    - email: MIME message parsing
    - quopri: Quoted-printable decoding
    - base64: Base64 decoding

Extracted Content
-----------------
Returns HtmlContent (same as HTML extractor):
    - content: Full text extracted from HTML
    - tables: Structured table data
    - headings: Document headings
    - links: Hyperlinks
    - metadata: Title, charset, etc.

Known Limitations
-----------------
- Embedded images are not extracted (only HTML text)
- JavaScript-generated content is not captured
- Some non-standard MHTML variants may not parse correctly
- Password-protected archives are not supported

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.mhtml_extractor import read_mhtml
    >>>
    >>> with open("archive.mhtml", "rb") as f:
    ...     for doc in read_mhtml(io.BytesIO(f.read()), path="archive.mhtml"):
    ...         print(f"Title: {doc.metadata.title}")
    ...         print(doc.content[:500])

See Also
--------
- RFC 2557: MIME Encapsulation of Aggregate Documents
- html_extractor: For HTML content processing

Maintenance Notes
-----------------
- Uses email.message_from_bytes for MIME parsing
- Handles both quoted-printable and base64 content transfer encodings
- Falls back to raw content if MIME parsing fails
"""

import base64
import io
import logging
import quopri
import re
from email import message_from_bytes
from email.message import Message
from typing import Any, Generator, Optional

from sharepoint2text.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.extractors.data_types import HtmlContent, HtmlMetadata
from sharepoint2text.extractors.html_extractor import read_html

logger = logging.getLogger(__name__)


def _decode_content(part: Message) -> bytes:
    """Decode MIME part content based on Content-Transfer-Encoding."""
    payload = part.get_payload(decode=False)

    if isinstance(payload, bytes):
        return payload

    if isinstance(payload, str):
        payload_bytes = payload.encode("utf-8", errors="replace")
    else:
        return b""

    encoding = part.get("Content-Transfer-Encoding", "").lower()

    if encoding == "quoted-printable":
        try:
            return quopri.decodestring(payload_bytes)
        except Exception:
            return payload_bytes
    elif encoding == "base64":
        try:
            # Remove whitespace that may be present in base64 content
            clean = re.sub(rb"\s+", b"", payload_bytes)
            return base64.b64decode(clean)
        except Exception:
            return payload_bytes
    else:
        # 7bit, 8bit, binary, or unspecified
        return payload_bytes


def _find_html_part(msg: Message) -> Optional[bytes]:
    """
    Find and extract the HTML content from an MHTML message.

    Searches through MIME parts to find text/html content.
    For multipart messages, checks all parts.
    """
    content_type = msg.get_content_type()

    # Direct HTML content
    if content_type == "text/html":
        return _decode_content(msg)

    # Multipart message - search parts
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                return _decode_content(part)

    # Check if the whole message might be HTML even without proper content-type
    # (some MHTML files have malformed headers)
    if not msg.is_multipart():
        content = _decode_content(msg)
        # Simple heuristic: check if it looks like HTML
        if b"<html" in content.lower() or b"<!doctype html" in content.lower():
            return content

    return None


def _extract_from_mhtml(content: bytes) -> Optional[bytes]:
    """
    Extract HTML content from MHTML bytes.

    Handles various MHTML formats and edge cases.
    """
    # Try standard MIME parsing
    try:
        msg = message_from_bytes(content)
        html_content = _find_html_part(msg)
        if html_content:
            return html_content
    except Exception as e:
        logger.debug("Standard MIME parsing failed: %s", e)

    # Fallback: Try to find HTML content directly in the file
    # Some MHTML files have non-standard structure
    try:
        # Look for Content-Type: text/html boundary
        html_start_patterns = [
            rb"Content-Type:\s*text/html[^\r\n]*\r?\n\r?\n",
            rb"Content-Type:\s*text/html[^\r\n]*\r?\n[^\r\n]*\r?\n\r?\n",
        ]

        for pattern in html_start_patterns:
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                start = match.end()
                # Find the next boundary or end of file
                boundary_match = re.search(rb"\r?\n--", content[start:])
                if boundary_match:
                    end = start + boundary_match.start()
                else:
                    end = len(content)

                html_bytes = content[start:end]

                # Check for Content-Transfer-Encoding before the HTML
                encoding_match = re.search(
                    rb"Content-Transfer-Encoding:\s*(\S+)",
                    content[max(0, match.start() - 200) : match.start()],
                    re.IGNORECASE,
                )
                if encoding_match:
                    encoding = encoding_match.group(1).lower()
                    if encoding == b"quoted-printable":
                        html_bytes = quopri.decodestring(html_bytes)
                    elif encoding == b"base64":
                        try:
                            html_bytes = base64.b64decode(
                                re.sub(rb"\s+", b"", html_bytes)
                            )
                        except Exception:
                            pass

                return html_bytes

        # Last resort: look for raw HTML in the content
        html_match = re.search(
            rb"(<html[^>]*>.*</html>)", content, re.IGNORECASE | re.DOTALL
        )
        if html_match:
            return html_match.group(1)

    except Exception as e:
        logger.debug("Fallback HTML extraction failed: %s", e)

    return None


def read_mhtml(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[HtmlContent, Any, None]:
    """
    Extract content from an MHTML (MIME HTML) web archive file.

    MHTML files are MIME-encoded archives containing HTML and embedded
    resources. This function extracts the HTML content and processes it
    using the standard HTML extractor.

    Args:
        file_like: BytesIO object containing the complete MHTML file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata in the returned HtmlContent.

    Yields:
        HtmlContent: Single HtmlContent object containing:
            - content: Extracted text from the HTML
            - tables: Structured table data
            - headings: Document headings (h1-h6)
            - links: Hyperlinks with text and href
            - metadata: Title, charset, description, etc.

    Raises:
        ExtractionFailedError: If extraction fails.

    Example:
        >>> import io
        >>> with open("archive.mhtml", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for doc in read_mhtml(data, path="archive.mhtml"):
        ...         print(f"Title: {doc.metadata.title}")
        ...         print(f"Content length: {len(doc.content)}")
    """
    try:
        file_like.seek(0)
        content = file_like.read()

        # Extract HTML from MHTML structure
        html_content = _extract_from_mhtml(content)

        if html_content is None:
            logger.warning("No HTML content found in MHTML file")
            metadata = HtmlMetadata()
            metadata.populate_from_path(path)
            yield HtmlContent(content="", metadata=metadata)
            return

        logger.debug("Extracted %d bytes of HTML from MHTML", len(html_content))

        # Use the HTML extractor to process the content
        html_buffer = io.BytesIO(html_content)
        for result in read_html(html_buffer, path=path):
            logger.info(
                "Extracted MHTML: %d characters, %d tables",
                len(result.content),
                len(result.tables),
            )
            yield result

    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract MHTML file", cause=exc) from exc
