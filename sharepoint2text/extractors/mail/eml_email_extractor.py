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
    >>> from sharepoint2text.extractors.mail.eml_email_extractor import (
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
import io
import logging
from typing import Any, Generator

from mailparser import parse_from_bytes

from sharepoint2text.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.extractors.data_types import (
    EmailAddress,
    EmailAttachment,
    EmailContent,
    EmailMetadata,
)
from sharepoint2text.mime_types import is_supported_mime_type

logger = logging.getLogger(__name__)


def _read_eml_format(payload: bytes) -> EmailContent:
    """
    Parse raw EML file bytes and construct an EmailContent object.

    This internal function performs the actual parsing work using mailparser,
    extracting headers, addresses, and body content into a structured format.

    Args:
        payload: Raw bytes of the EML file content. Should be the complete
            file contents, including all headers and body parts.

    Returns:
        EmailContent: Populated dataclass with all extracted email data.

    Raises:
        IndexError: If the From header is missing or malformed. The mailparser
            library returns from_ as a list of (name, address) tuples.
        AttributeError: If mailparser fails to parse the payload structure.

    Implementation Notes:
        - mailparser.from_ returns a list; we take the first entry as the
          sender (multiple From addresses are rare and non-standard)
        - CC, BCC, and Reply-To fields filter out empty/malformed entries
        - Date is converted to ISO format string for consistency
        - Both text/plain and text/html bodies are extracted if present
        - Body content may be a list (multipart) or string (single part)

    Maintenance Considerations:
        The mailparser library version may affect tuple structure. Current
        implementation expects (name, address) tuples. Verify after upgrades.
    """
    mail = parse_from_bytes(payload)

    # Extract sender address - mailparser returns list of (name, address) tuples
    # First entry is the primary From address
    from_email = EmailAddress(mail.from_[0][0], mail.from_[0][1])

    # Extract recipient lists - filter malformed entries that lack address
    to_emails = [EmailAddress(name=t[0], address=t[1]) for t in mail.to]

    cc = [
        EmailAddress(name=t[0], address=t[1])
        for t in mail.cc
        if t and len(t) > 1 and t[1]
    ]
    bcc = [
        EmailAddress(name=t[0], address=t[1])
        for t in mail.bcc
        if t and len(t) > 1 and t[1]
    ]
    reply_to = [
        EmailAddress(name=t[0], address=t[1])
        for t in mail.reply_to
        if t and len(t) > 1 and t[1]
    ]

    # Extract and format date as ISO string for consistent representation
    date_str = ""
    if mail.date:
        date_str = mail.date.isoformat()

    metadata = EmailMetadata(
        date=date_str,
        message_id=mail.message_id or "",
    )

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

    attachments: list[EmailAttachment] = []
    for attachment in mail.attachments:
        filename = attachment.get("filename") or "attachment"
        mime_type = attachment.get("mail_content_type") or "application/octet-stream"
        payload = attachment.get("payload") or b""
        is_binary = bool(attachment.get("binary"))

        if is_binary:
            if isinstance(payload, str):
                data = base64.b64decode(payload)
            else:
                data = base64.b64decode(payload)
        else:
            if isinstance(payload, str):
                data = payload.encode("utf-8", errors="ignore")
            else:
                data = payload

        data_stream = io.BytesIO(data)
        data_stream.seek(0)
        attachments.append(
            EmailAttachment(
                filename=filename,
                mime_type=mime_type,
                data=data_stream,
                is_supported_mime_type=is_supported_mime_type(mime_type),
            )
        )

    return EmailContent(
        subject=mail.subject or "",
        from_email=from_email,
        to_emails=to_emails,
        to_cc=cc,
        to_bcc=bcc,
        reply_to=reply_to,
        in_reply_to=mail.in_reply_to or "",
        body_plain=body_plain,
        body_html=body_html,
        attachments=attachments,
        metadata=metadata,
    )


def read_eml_format_mail(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    """
    Read an EML file and extract its content as EmailContent.

    Primary entry point for EML file extraction. Accepts a BytesIO object
    containing the raw email data and yields EmailContent objects.

    This function uses a generator pattern for API consistency with other
    email extractors (mbox can contain multiple emails), even though EML
    files contain exactly one email.

    Args:
        file_like: BytesIO object containing the complete EML file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned EmailContent.metadata object. Useful for tracking
            source files in batch processing scenarios.

    Yields:
        EmailContent: Single EmailContent object containing all extracted
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
        - Entire file is loaded into memory via getvalue()
        - For very large emails (>10MB), consider memory implications
        - Stream position is reset; original position is not preserved
    """
    try:
        file_like.seek(0)
        content = _read_eml_format(file_like.getvalue())

        if path:
            content.metadata.populate_from_path(path)

        logger.info("Extracted EML")

        yield content
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract EML file", cause=exc) from exc
