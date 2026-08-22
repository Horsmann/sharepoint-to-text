"""
MSG Email Extractor
===================

Extracts text content and metadata from Microsoft Outlook .msg files.

This module handles the proprietary MSG format used by Microsoft Outlook
for saving individual email messages. MSG files use the OLE (Object Linking
and Embedding) Compound File format, also known as Microsoft Compound
Document format.

File Format Background
----------------------
The MSG format is a binary format based on the Compound File Binary Format
(CFBF), which is the same underlying structure used by older Microsoft
Office formats (.doc, .xls, .ppt). Key characteristics:

    - Binary format (not human-readable)
    - Based on OLE/COM structured storage
    - Contains message properties as MAPI properties
    - Attachments stored as nested storage objects
    - Rich text and HTML body variants may be present

MSG files are commonly encountered when:
    - Users drag emails from Outlook to save them
    - Email archiving systems export from Exchange
    - Legal/compliance exports from Microsoft 365
    - SharePoint document libraries with saved emails

Dependencies
------------
msg_parser: https://github.com/vikramarsid/msg_parser
    pip install msg_parser

    Provides:
    - OLE compound document parsing
    - MAPI property extraction
    - Body text retrieval (plain, RTF, HTML)
    - Attachment enumeration (extracted as binary)

The msg_parser library uses the olefile library internally for
OLE parsing.

Known Limitations
-----------------
- RTF body is not separately extracted (uses same body for both plain/HTML)
- Embedded images in HTML are not processed
- Some MAPI properties may not be extracted (extended properties)
- Calendar items and meeting requests may have incomplete data
- Encrypted/protected messages will fail to parse

MSG Format Quirks
-----------------
Microsoft's MSG format has several peculiarities that affect parsing:

1. Recipient Format:
   Recipients may be stored in various formats depending on Outlook version:
   - "Display Name <email@example.com>" (standard)
   - Just the email address
   - Just the display name (for internal Exchange recipients)
   - X.500 distinguished names (legacy Exchange)

2. Date Handling:
   Dates are stored as MAPI PT_SYSTIME properties and converted to
   RFC 2822 format strings by msg_parser. Some messages may have
   malformed or missing date properties.

3. Body Content:
   MSG files may contain body in three formats:
   - Plain text (PidTagBody)
   - RTF (PidTagRtfCompressed) - compressed RTF
   - HTML (PidTagBodyHtml)

   The msg_parser library provides .body which typically contains
   plain text. HTML may require additional extraction.

4. Character Encoding:
   Modern MSG files use Unicode (UTF-16LE for strings), but older
   files may use code pages. The msg_parser library handles this.

Usage
-----
    >>> import io
    >>> from sharepoint2text.parsing.extractors.mail.msg_email_extractor import (
    ...     read_msg_format_mail
    ... )
    >>>
    >>> with open("message.msg", "rb") as f:
    ...     msg_data = io.BytesIO(f.read())
    ...     for email in read_msg_format_mail(msg_data, path="message.msg"):
    ...         print(f"Subject: {email.subject}")
    ...         print(f"From: {email.from_email.address}")
    ...         print(f"Body preview: {email.body_plain[:100]}...")

See Also
--------
- eml_email_extractor: For RFC 5322 text format emails
- mbox_email_extractor: For Unix mailbox format
"""

import io
import logging
import re
from email.utils import parsedate_to_datetime
from typing import Any, Generator

from msg_parser import MsOxMessage  # type: ignore[import-untyped]
from olefile import OleFileIO  # type: ignore[import-untyped]

from sharepoint2text.parsing.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.parsing.extractors._legacy_types import (
    EmailAddress,
    EmailAttachment,
    EmailContent,
    EmailMetadata,
)
from sharepoint2text.parsing.extractors.html_extractor import (
    _HtmlTextExtractor,
    _HtmlTreeBuilder,
)
from sharepoint2text.parsing.mime_types import is_supported_mime_type

logger = logging.getLogger(__name__)

_HTML_HINT_RE = re.compile(
    r"<(html|head|body|p|div|br|span|table|tr|td|style|script)(\\s|>)",
    re.IGNORECASE,
)


def _looks_like_html(text: str) -> bool:
    if not text:
        return False
    lowered = text.lstrip().lower()
    if lowered.startswith("<!doctype") or "<html" in lowered or "<body" in lowered:
        return True
    return _HTML_HINT_RE.search(text) is not None


def _html_to_text(html_text: str) -> str:
    try:
        parser = _HtmlTreeBuilder()
        parser.feed(html_text)
        root = parser.get_tree()
        extractor = _HtmlTextExtractor(root)
        text = extractor.extract()
        text = text.replace("\u200b", "").replace("\ufeff", "")
        return text.strip()
    except (ValueError, RecursionError):
        return html_text.strip()


def _read_ole_string(ole: OleFileIO, storage: str, stream_name: str) -> str:
    try:
        raw = ole.openstream([storage, stream_name]).read()
    except (OSError, IOError):
        return ""
    try:
        raw_str: str = raw.decode("utf-16-le", errors="ignore").rstrip("\x00")
        return raw_str
    except (UnicodeDecodeError, ValueError):
        return ""


def _attachment_data_to_bytes(value: Any) -> bytes:
    """Normalize msg-parser attachment payloads to bytes."""
    if value is None:
        return b""
    if isinstance(value, bytes):
        return value
    if isinstance(value, bytearray):
        return bytes(value)
    if isinstance(value, memoryview):
        return value.tobytes()
    if isinstance(value, str):
        return value.encode("utf-8", errors="ignore")
    return bytes(value)


def _extract_msg_attachments_from_msg(msg: MsOxMessage) -> list[EmailAttachment]:
    """Extract attachments from msg_parser output without re-reading the OLE file."""
    attachments: list[EmailAttachment] = []
    msg_attachments = getattr(msg, "attachments", None) or []

    for index, attachment in enumerate(msg_attachments, start=1):
        data = _attachment_data_to_bytes(getattr(attachment, "data", b""))
        filename = (
            getattr(attachment, "Filename", None)
            or getattr(attachment, "AttachLongFilename", None)
            or getattr(attachment, "AttachFilename", None)
            or f"attachment-{index}"
        )
        mime_type = (
            getattr(attachment, "AttachMimeTag", None) or "application/octet-stream"
        )

        attachments.append(
            EmailAttachment(
                filename=filename,
                mime_type=mime_type,
                data=io.BytesIO(data),
                is_supported_mime_type=is_supported_mime_type(mime_type),
            )
        )

    return attachments


def _extract_msg_attachments_from_ole(file_like: io.BytesIO) -> list[EmailAttachment]:
    attachments: list[EmailAttachment] = []

    file_like.seek(0)
    with OleFileIO(file_like) as ole:
        storages = ole.listdir(streams=False, storages=True)
        attach_storages = [
            storage[0]
            for storage in storages
            if len(storage) == 1 and storage[0].startswith("__attach_version1.0_")
        ]

        for index, storage in enumerate(attach_storages, start=1):
            try:
                data = ole.openstream([storage, "__substg1.0_37010102"]).read()
            except (OSError, IOError):
                continue

            filename = (
                _read_ole_string(ole, storage, "__substg1.0_3707001F")
                or _read_ole_string(ole, storage, "__substg1.0_3704001F")
                or f"attachment-{index}"
            )
            mime_type = (
                _read_ole_string(ole, storage, "__substg1.0_370E001F")
                or "application/octet-stream"
            )

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

    return attachments


def _parse_single_recipient(raw: str) -> EmailAddress | None:
    """
    Parse a single recipient string into an EmailAddress object.

    Handles various recipient formats found in MSG files:
        - "Display Name <email@example.com>" - standard format
        - "<email@example.com>" - angle brackets only
        - "email@example.com" - raw email address
        - "Display Name" - name only (common for Exchange internal recipients)

    Args:
        raw: Single recipient string to parse. May include display name,
            email address, or both.

    Returns:
        EmailAddress object with parsed name and address fields, or None
        if the input is empty or whitespace-only.

    Examples:
        >>> _parse_single_recipient("John Doe <john@example.com>")
        EmailAddress(name='John Doe', address='john@example.com')
        >>> _parse_single_recipient("<admin@example.com>")
        EmailAddress(name='', address='admin@example.com')
        >>> _parse_single_recipient("user@example.com")
        EmailAddress(name='', address='user@example.com')
        >>> _parse_single_recipient("John Doe")
        EmailAddress(name='John Doe', address='')
        >>> _parse_single_recipient("")
        None

    Implementation Notes:
        Uses regex to find angle-bracketed email first, then falls back
        to checking for @ symbol to distinguish bare emails from names.
        Quotes around display names are stripped.
    """
    raw = raw.strip()
    if not raw:
        return None

    # Look for <email> pattern at the end of the string
    match = re.search(r"<([^>]+)>\s*$", raw)
    if match:
        address = match.group(1).strip()
        # Name is everything before the angle brackets, strip quotes
        name = raw[: match.start()].strip().strip("\"'")
        return EmailAddress(name=name, address=address)

    # No angle brackets - check if it's just an email address
    # Email addresses have @ and no spaces
    if "@" in raw and " " not in raw:
        return EmailAddress(name="", address=raw)

    # No email found - treat entire string as display name only
    # This is common for Exchange internal recipients
    return EmailAddress(name=raw, address="")


def _parse_multi_recipients(raw: str | list[str]) -> list[EmailAddress]:
    """
    Parse recipient string(s) that may contain multiple addresses.

    MSG files may store recipients as either a list of strings or a
    single semicolon/comma-separated string. This function handles both.

    Args:
        raw: Either a list of recipient strings, or a single string
            containing multiple recipients separated by semicolon or comma.
            Examples:
            - ['John Doe <john@example.com>']
            - 'Alice <alice@x.com>; Bob <bob@x.com>'
            - 'user1@x.com, user2@x.com'

    Returns:
        List of EmailAddress objects. Empty list if input is empty/None.
        Only entries with a name or address are included.

    Examples:
        >>> _parse_multi_recipients(['User <user@x.com>'])
        [EmailAddress(name='User', address='user@x.com')]
        >>> _parse_multi_recipients('A <a@x.com>; B <b@x.com>')
        [EmailAddress(name='A', address='a@x.com'),
         EmailAddress(name='B', address='b@x.com')]

    Implementation Notes:
        - Recursively handles list input by processing each element
        - Splits strings on both semicolon and comma (common delimiters)
        - Filters out entries that have neither name nor address
    """
    if not raw:
        return []

    # Handle list input by recursively processing each item
    if isinstance(raw, list):
        results = []
        for item in raw:
            results.extend(_parse_multi_recipients(item))
        return results

    # Split by semicolon or comma (common recipient separators in Outlook)
    parts = re.split(r"[;,]", raw)

    recipients = []
    for part in parts:
        addr = _parse_single_recipient(part)
        # Only include if we got either a name or an address
        if addr and (addr.name or addr.address):
            recipients.append(addr)

    return recipients


def read_msg_format_mail(
    file_like: io.BytesIO,
    path: str | None = None,
    *,
    ignore_images: bool = False,
    include_attachments: bool = True,
) -> Generator[EmailContent, Any, None]:
    """
    Read a Microsoft Outlook MSG file and extract its content.

    Primary entry point for MSG file extraction. Parses the OLE compound
    document structure and extracts email headers, addresses, body
    content, and attachments into an EmailContent object.

    This function uses a generator pattern for API consistency with other
    email extractors, even though MSG files contain exactly one email.

    Args:
        file_like: BytesIO object containing the complete MSG file data.
            Must be positioned at the start of the file. The entire MSG
            binary content should be readable from this stream.
        path: Optional filesystem path to source file. If provided, populates
            file metadata (filename, extension, folder) in the returned
            EmailContent.metadata. Useful for batch processing scenarios.
        ignore_images: If True, skip image extraction (not applicable for this format).

    Yields:
        EmailContent: Single EmailContent object containing all extracted
            data. The generator yields exactly one item for valid MSG files.
            Attachments are stored on EmailContent.attachments.

    Raises:
        Exception: Various exceptions from msg_parser for:
            - Corrupted or invalid OLE structure
            - Missing required MAPI properties
            - Encrypted/protected messages

    Example:
        >>> import io
        >>> with open("meeting.msg", "rb") as f:
        ...     for email in read_msg_format_mail(io.BytesIO(f.read())):
        ...         print(f"Subject: {email.subject}")
        ...         print(f"From: {email.from_email.name} <{email.from_email.address}>")
        ...         print(f"Date: {email.metadata.date}")

    Implementation Notes:
        - msg_parser.MsOxMessage handles OLE parsing internally
        - Sender is extracted via sender property and parsed as recipient list
        - To, Cc, Bcc fields may be strings or lists depending on msg_parser version
        - reply_to is normalized to List[EmailAddress]
        - body_plain is plain text; HTML is converted when detected
        - Attachments are returned as EmailAttachment dataclasses
        - For HTML body, additional extraction from RTF may be needed

    Maintenance Considerations:
        - msg_parser API may change between versions; verify property names
        - Some MSG files may have missing properties (sender, date, etc.)
        - HTML body extraction could be enhanced by parsing RTF content

    Known Issues:
        - body_html uses the raw HTML body when detected
    """
    source_path = path or "<in-memory>"
    logger.info("Entering MSG extraction: %s", source_path)
    try:
        file_like.seek(0)
        msg = MsOxMessage(file_like)

        # Build metadata with date and message ID
        sent_date = ""
        if msg.sent_date:
            try:
                sent_date = parsedate_to_datetime(msg.sent_date).isoformat()
            except (TypeError, ValueError):
                logger.debug("Unable to parse MSG sent_date: %r", msg.sent_date)

        meta = EmailMetadata(
            message_id=msg.message_id or "",
            date=sent_date,
        )

        # Parse sender - expecting at least one result
        # msg.sender may be a single string or list depending on version
        sender_list = _parse_multi_recipients(msg.sender)
        from_email = sender_list[0] if sender_list else EmailAddress()

        attachments = []
        msg_attachments = (
            getattr(msg, "attachments", None) or [] if include_attachments else []
        )
        if msg_attachments:
            # Prefer raw OLE attachment payloads for byte-level fidelity.
            attachments = _extract_msg_attachments_from_ole(file_like)
            if not attachments:
                attachments = _extract_msg_attachments_from_msg(msg)

        raw_body = msg.body or ""
        if _looks_like_html(raw_body):
            body_plain = _html_to_text(raw_body)
            body_html = raw_body
        else:
            body_plain = raw_body
            body_html = ""

        content = EmailContent(
            subject=msg.subject or "",
            from_email=from_email,
            to_emails=_parse_multi_recipients(msg.to),
            to_cc=_parse_multi_recipients(msg.cc),
            to_bcc=_parse_multi_recipients(msg.bcc),
            reply_to=_parse_multi_recipients(msg.reply_to),
            body_plain=body_plain,
            body_html=body_html,
            attachments=attachments,
            metadata=meta,
        )

        if path:
            content.metadata.populate_from_path(path)

        yield content
    except ExtractionError:
        raise
    except (ValueError, TypeError, AttributeError, OSError, UnicodeDecodeError) as exc:
        raise ExtractionFailedError("Failed to extract MSG file", cause=exc) from exc
    finally:
        logger.info("Leaving MSG extraction: %s", source_path)
