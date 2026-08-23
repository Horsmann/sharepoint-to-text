"""
EML Email Extractor
===================

Extracts text content and metadata from .eml files (RFC 5322 / MIME format).

This module handles standard Internet Message Format emails, which are the
de facto standard for email interchange. EML files are plain text files
containing email headers and body content, potentially with MIME multipart
structures for attachments and alternative content types.

File Format Background
----------------------
EML format follows RFC 5322 (Internet Message Format) and RFC 2045-2049 (MIME).
Files typically contain:
    - Headers (From, To, Subject, Date, Message-ID, etc.)
    - Body content (plain text, HTML, or both via multipart/alternative)
    - Attachments (via multipart/mixed, extracted by this module)

The format is text-based and human-readable, making it widely compatible
but potentially large for emails with encoded attachments.

Dependencies
------------
mailparser: https://github.com/SpamScope/mail-parser
    pip install mail-parser

    Provides robust MIME parsing with automatic handling of:
    - Character encoding detection and conversion
    - Multipart message structure navigation
    - Header decoding (RFC 2047 encoded words)
    - Date parsing to datetime objects

Known Limitations
-----------------
- Embedded images in HTML are not processed
- Malformed headers may cause partial extraction
- Very large emails may consume significant memory (entire file loaded)

Encoding Considerations
-----------------------
The mailparser library handles most encoding scenarios, but edge cases exist:
- Emails with incorrect charset declarations
- Mixed encodings within a single message
- Legacy 8-bit headers (non-RFC compliant)

For problematic emails, body content may contain replacement characters
where decoding failed.

Usage
-----
    >>> import io
    >>> from sharepoint2text.parsing.extractors.mail.eml_email_extractor import (
    ...     read_eml_format_mail
    ... )
    >>>
    >>> with open("message.eml", "rb") as f:
    ...     for email in read_eml_format_mail(io.BytesIO(f.read())):
    ...         print(f"From: {email.from_email.address}")
    ...         print(f"Subject: {email.subject}")
    ...         print(f"Body: {email.body_plain[:100]}...")

See Also
--------
- mbox_email_extractor: For Unix mailbox format (multiple emails)
- msg_email_extractor: For Microsoft Outlook .msg format
"""

import base64
import email.utils
import io
import logging
from typing import Any, Generator

from mailparser import parse_from_bytes  # type: ignore[import-untyped]

from sharepoint2text.parsing.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.parsing.extractors._model import source_metadata
from sharepoint2text.parsing.extractors.mail._model import Address, build_email_document
from sharepoint2text.parsing.mime_types import is_supported_mime_type
from sharepoint2text.parsing.models import Attachment, ExtractedDocument

logger = logging.getLogger(__name__)


def _to_email_address(raw: Any) -> Address | None:
    """Best-effort conversion of a mailparser address value to EmailAddress."""
    if raw is None:
        return None

    if isinstance(raw, (tuple, list)):
        name = str(raw[0]).strip() if len(raw) > 0 and raw[0] is not None else ""
        address = str(raw[1]).strip() if len(raw) > 1 and raw[1] is not None else ""
        if not address and len(raw) == 1 and name:
            parsed_name, parsed_addr = email.utils.parseaddr(name)
            if parsed_addr:
                name, address = parsed_name.strip(), parsed_addr.strip()
        if not (name or address):
            return None
        return name, address

    text = str(raw).strip()
    if not text:
        return None
    parsed_name, parsed_addr = email.utils.parseaddr(text)
    if parsed_addr:
        return parsed_name.strip(), parsed_addr.strip()
    return text, ""


def _parse_mailparser_address_list(raw: Any) -> list[Address]:
    """Normalize mailparser address collections to List[EmailAddress]."""
    if not raw:
        return []

    if isinstance(raw, (list, tuple)):
        result = []
        for entry in raw:
            addr = _to_email_address(entry)
            if addr is not None and (addr[0] or addr[1]):
                result.append(addr)
        return result

    addr = _to_email_address(raw)
    if addr is None or not (addr[0] or addr[1]):
        return []
    return [addr]


def _read_eml_format(
    payload: bytes, *, include_attachments: bool = True
) -> ExtractedDocument:
    """
    Parse raw EML bytes and construct a canonical document.

    This internal function performs the actual parsing work using mailparser,
    extracting headers, addresses, and body content into a structured format.

    Args:
        payload: Raw bytes of the EML file content. Should be the complete
            file contents, including all headers and body parts.

    Returns:
        Canonical document with all extracted email data.

    Implementation Notes:
        - mailparser.from_ is normalized to a best-effort EmailAddress
          (missing/malformed sender headers become empty fields)
        - CC, BCC, and Reply-To fields filter out empty/malformed entries
        - Date is converted to ISO format string for consistency
        - Both text/plain and text/html bodies are extracted if present
        - Body content may be a list (multipart) or string (single part)

    Maintenance Considerations:
        The mailparser library version may affect tuple structure. Current
        implementation expects (name, address) tuples. Verify after upgrades.
    """
    mail = parse_from_bytes(payload)

    # Extract sender address - tolerate missing/malformed values in real-world data.
    from_candidates = _parse_mailparser_address_list(getattr(mail, "from_", None))
    from_email = from_candidates[0] if from_candidates else ("", "")

    # Extract recipient lists with robust normalization.
    to_emails = _parse_mailparser_address_list(getattr(mail, "to", None))
    cc = _parse_mailparser_address_list(getattr(mail, "cc", None))
    bcc = _parse_mailparser_address_list(getattr(mail, "bcc", None))
    reply_to = _parse_mailparser_address_list(getattr(mail, "reply_to", None))

    # Extract and format date as ISO string for consistent representation
    date_str = ""
    if mail.date:
        try:
            date_str = mail.date.isoformat()
        except AttributeError:
            date_str = str(mail.date)

    # Body extraction - mailparser uses text_plain/text_html attributes
    # These may be lists (multipart) or single strings depending on structure
    body_plain = ""
    if mail.text_plain:
        if isinstance(mail.text_plain, list):
            body_plain = "\n".join(mail.text_plain)
        else:
            body_plain = str(mail.text_plain)

    body_html = ""
    if mail.text_html:
        if isinstance(mail.text_html, list):
            body_html = "\n".join(mail.text_html)
        else:
            body_html = str(mail.text_html)

    attachments: list[Attachment] = []
    if include_attachments:
        for attachment in getattr(mail, "attachments", None) or []:
            filename = attachment.get("filename") or "attachment"
            mime_type = (
                attachment.get("mail_content_type") or "application/octet-stream"
            )
            payload = attachment.get("payload") or b""
            is_binary = bool(attachment.get("binary"))

            if is_binary:
                try:
                    data = base64.b64decode(payload)
                except (ValueError, TypeError):
                    logger.debug(
                        "Unable to base64-decode EML attachment payload: %s", filename
                    )
                    data = b""
            else:
                if isinstance(payload, str):
                    data = payload.encode("utf-8", errors="ignore")
                else:
                    data = payload

            attachments.append(
                Attachment(
                    filename=filename,
                    media_type=mime_type,
                    data=data,
                    properties={
                        "email.is_supported_mime_type": is_supported_mime_type(
                            mime_type
                        )
                    },
                )
            )

    return build_email_document(
        source_format="eml",
        path=None,
        subject=mail.subject or "",
        sender=from_email,
        recipients=to_emails,
        cc=cc,
        bcc=bcc,
        reply_to=reply_to,
        in_reply_to=mail.in_reply_to or "",
        body_plain=body_plain,
        body_html=body_html,
        attachments=attachments,
        date=date_str,
        message_id=mail.message_id or "",
    )


def read_eml_format_mail(
    file_like: io.BytesIO,
    path: str | None = None,
    *,
    ignore_images: bool = False,
    include_attachments: bool = True,
) -> Generator[ExtractedDocument, Any, None]:
    """
    Read an EML file and extract its content as a canonical document.

    Primary entry point for EML file extraction. Accepts a BytesIO object
    containing the raw email data and yields canonical documents.

    This function uses a generator pattern for API consistency with other
    email extractors (mbox can contain multiple emails), even though EML
    files contain exactly one email.

    Args:
        file_like: BytesIO object containing the complete EML file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned document source metadata. Useful for tracking
            source files in batch processing scenarios.
        ignore_images: If True, skip image extraction (not applicable for this format).

    Yields:
        ExtractedDocument: Single canonical document containing all extracted
            data. The generator will yield exactly one item for valid EML
            files.

    Raises:
        Exception: Various exceptions may propagate from mailparser for
            malformed or corrupted EML files. Common issues include:
            - Missing required headers (From, Date)
            - Invalid MIME structure
            - Encoding errors in binary payloads

    Example:
        >>> import io
        >>> eml_data = b"From: sender@example.com\\r\\nTo: recipient@example.com..."
        >>> buffer = io.BytesIO(eml_data)
        >>> for email in read_eml_format_mail(buffer, path="/archive/msg.eml"):
        ...     print(email.subject)
        ...     print(email.metadata.filename)  # "msg.eml"

    Performance Notes:
        - Parses directly from the input stream to avoid an additional full-copy
          buffer in memory.
    """
    try:
        file_like.seek(0)
        content = _read_eml_format(
            file_like.read(), include_attachments=include_attachments
        )

        content.source = source_metadata(path)

        logger.debug(
            "Extracted EML: attachments=%d",
            len(content.attachments),
        )

        yield content
    except ExtractionError:
        raise
    except (
        IndexError,
        AttributeError,
        ValueError,
        TypeError,
        UnicodeDecodeError,
    ) as exc:
        raise ExtractionFailedError("Failed to extract EML file", cause=exc) from exc
