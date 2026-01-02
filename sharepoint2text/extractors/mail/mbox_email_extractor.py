"""
MBOX Email Extractor
====================

Extracts text content and metadata from Unix mailbox (.mbox) files.

This module handles the mbox format, which stores multiple email messages
concatenated in a single file. Each message is preceded by a "From " line
(note the space) that serves as a message delimiter.

File Format Background
----------------------
The mbox format originated in Unix systems and remains widely used for
email archiving and migration. Key characteristics:

    - Plain text format, human-readable
    - Multiple messages in a single file
    - Messages separated by "From " lines at the start of each email
    - Each message follows RFC 5322 format (same as EML)
    - No standardized escape mechanism (various mbox variants exist)

Common mbox sources include:
    - Unix mail clients (mutt, pine)
    - Email exports from Thunderbird, Apple Mail
    - Google Takeout Gmail exports
    - Legacy mail server archives

Dependencies
------------
Python Standard Library:
    - mailbox: Core mbox parsing and iteration
    - email: RFC 5322 message parsing
    - email.header: Encoded header word handling (RFC 2047)
    - email.utils: Address parsing and date handling
    - tempfile: Temporary file creation for mbox processing

No external dependencies required.

Implementation Details
----------------------
The Python mailbox module requires a filesystem path, not a file object.
This necessitates writing the BytesIO contents to a temporary file before
parsing. The temporary file is cleaned up after processing.

Known Limitations
-----------------
- Requires temporary file creation (disk I/O overhead)
- Attachments are not extracted (only body text)
- "From " lines within message bodies may cause message boundary issues
  in poorly-formed mbox files (the "mboxrd" escaping is not handled)
- Large mbox files may be slow due to sequential processing
- Memory usage scales with individual message size, not file size

Encoding Handling
-----------------
This module uses Python's email library for encoding handling:
- Headers: RFC 2047 encoded words are decoded automatically
- Bodies: Charset from Content-Type header, fallback to UTF-8
- Errors: Replacement characters used for undecodable bytes

Common encoding issues:
- Legacy emails with CP1252/ISO-8859-1 marked as ASCII
- Mixed encodings within MIME parts
- Missing or incorrect charset declarations

The decode_header_value() function handles these gracefully with fallbacks.

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.mail.mbox_email_extractor import (
    ...     read_mbox_format_mail
    ... )
    >>>
    >>> with open("archive.mbox", "rb") as f:
    ...     mbox_data = io.BytesIO(f.read())
    ...     for email in read_mbox_format_mail(mbox_data, path="archive.mbox"):
    ...         print(f"Subject: {email.subject}")
    ...         print(f"From: {email.from_email.address}")
    ...         print("---")

See Also
--------
- eml_email_extractor: For single RFC 5322 messages
- msg_email_extractor: For Microsoft Outlook .msg format
"""

import email
import email.header
import email.message
import email.utils
import io
import logging
import re
from email.utils import parsedate_to_datetime
from typing import Any, Generator

from sharepoint2text.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.extractors.data_types import (
    EmailAddress,
    EmailContent,
    EmailMetadata,
)

logger = logging.getLogger(__name__)

# Pattern to match mbox "From " separator lines
# Format: "From sender@example.com Mon Jan  1 00:00:00 2024"
# The line must start with "From " followed by an address and a date
MBOX_FROM_PATTERN = re.compile(rb"^From \S+.*\d{4}\r?\n", re.MULTILINE)


def _split_mbox_messages(data: bytes) -> list[bytes]:
    """
    Split mbox data into individual message bytes without using temp files.

    The mbox format separates messages with "From " lines at the start of a line.
    This function finds all such separators and splits the data accordingly.

    Args:
        data: Raw bytes of the entire mbox file.

    Returns:
        List of bytes, each containing a single email message.
        Empty list if no valid messages found.

    Implementation Notes:
        - Uses regex to find "From " separator lines
        - Each message starts after a "From " line
        - The "From " line itself is NOT part of the message content
        - Handles both Unix (LF) and Windows (CRLF) line endings
    """
    if not data:
        return []

    messages = []

    # Find all "From " line positions
    matches = list(MBOX_FROM_PATTERN.finditer(data))

    if not matches:
        # No "From " lines found - might be a single message without separator
        # or not a valid mbox format
        return []

    for i, match in enumerate(matches):
        # Message content starts after the "From " line
        msg_start = match.end()

        # Message ends at the next "From " line or end of data
        if i + 1 < len(matches):
            msg_end = matches[i + 1].start()
        else:
            msg_end = len(data)

        # Extract message bytes (strip trailing blank lines between messages)
        msg_bytes = data[msg_start:msg_end].rstrip(b"\r\n")

        if msg_bytes:
            messages.append(msg_bytes)

    return messages


def decode_header_value(value: str | None) -> str:
    """
    Decode an email header value, handling RFC 2047 encoded words.

    Email headers may contain encoded words for non-ASCII characters,
    formatted as: =?charset?encoding?encoded_text?=

    This function decodes these encoded words and handles various
    edge cases and malformed encodings.

    Args:
        value: Raw header value string, potentially containing encoded
            words. May be None for missing headers.

    Returns:
        Decoded string with all encoded words converted to Unicode.
        Returns empty string if value is None or empty.

    Examples:
        >>> decode_header_value("=?utf-8?B?SGVsbG8gV29ybGQ=?=")
        'Hello World'
        >>> decode_header_value("Plain ASCII header")
        'Plain ASCII header'
        >>> decode_header_value(None)
        ''

    Implementation Notes:
        - email.header.decode_header returns list of (bytes/str, charset) tuples
        - If charset is None, the part was not encoded
        - LookupError catches unknown charset names
        - UnicodeDecodeError catches invalid byte sequences
        - Fallback to UTF-8 with replacement handles all edge cases
    """
    if not value:
        return ""

    decoded_parts = []
    for part, charset in email.header.decode_header(value):
        if isinstance(part, bytes):
            # Encoded word - use declared charset or fall back to UTF-8
            charset = charset or "utf-8"
            try:
                decoded_parts.append(part.decode(charset, errors="replace"))
            except (LookupError, UnicodeDecodeError):
                # Unknown charset or decode error - force UTF-8
                decoded_parts.append(part.decode("utf-8", errors="replace"))
        else:
            # Already a string (unencoded ASCII)
            decoded_parts.append(part)

    return "".join(decoded_parts)


def parse_email_address(addr_string: str | None) -> EmailAddress:
    """
    Parse a single email address string into an EmailAddress object.

    Handles standard RFC 5322 address formats:
        - "Display Name <email@example.com>"
        - "<email@example.com>"
        - "email@example.com"

    Args:
        addr_string: Raw address string from email header.
            May be None for missing headers.

    Returns:
        EmailAddress: Parsed address with name and address fields.
            Returns EmailAddress with empty fields if input is None.

    Examples:
        >>> parse_email_address("John Doe <john@example.com>")
        EmailAddress(name='John Doe', address='john@example.com')
        >>> parse_email_address("<admin@example.com>")
        EmailAddress(name='', address='admin@example.com')
        >>> parse_email_address(None)
        EmailAddress(name='', address='')

    Notes:
        - Display names may contain RFC 2047 encoded words, which are decoded
        - Malformed addresses return empty fields rather than raising exceptions
    """
    if not addr_string:
        return EmailAddress()

    name, address = email.utils.parseaddr(addr_string)
    return EmailAddress(name=decode_header_value(name), address=address)


def parse_email_addresses(addr_string: str | None) -> list[EmailAddress]:
    """
    Parse a comma-separated list of email addresses.

    Handles the To, Cc, Bcc header formats which may contain multiple
    addresses separated by commas or semicolons.

    Args:
        addr_string: Raw address list string from email header.
            Example: "Alice <alice@x.com>, Bob <bob@x.com>"

    Returns:
        List of EmailAddress objects. Empty list if input is None or empty.
        Only addresses with valid email portions are included.

    Examples:
        >>> addrs = parse_email_addresses("A <a@x.com>, B <b@x.com>")
        >>> len(addrs)
        2
        >>> addrs[0].name
        'A'

    Notes:
        - email.utils.getaddresses handles complex quoting and escaping
        - Entries without an address portion are filtered out
        - Display names are decoded for RFC 2047 encoded words
    """
    if not addr_string:
        return []

    addresses = email.utils.getaddresses([addr_string])
    result = []
    for name, addr in addresses:
        if addr:  # Only include if there's an actual address
            result.append(EmailAddress(name=decode_header_value(name), address=addr))
    return result


def get_body_content(message: email.message.Message) -> tuple[str, str]:
    """
    Extract plain text and HTML body content from an email message.

    Navigates MIME multipart structure to find text/plain and text/html
    parts. For multipart messages, walks through all parts looking for
    body content. Attachments are skipped.

    Args:
        message: Parsed email.message.Message object.

    Returns:
        Tuple of (plain_text_body, html_body). Either may be empty string
        if that content type is not present in the message.

    Implementation Details:
        - For multipart messages, walks all parts recursively
        - Content-Disposition: attachment parts are skipped
        - Takes first text/plain and first text/html found
        - Charset from Content-Type header used for decoding
        - Unknown charsets fall back to UTF-8 with replacement

    MIME Structure Handling:
        - multipart/alternative: Contains both plain and HTML versions
        - multipart/mixed: Contains body plus attachments
        - multipart/related: HTML with inline images
        - Non-multipart: Single content type message

    Performance Notes:
        - Decodes body content from bytes to string
        - Large attachments are read but discarded if they're text/*
        - Consider adding size limits for production use
    """
    body_plain = ""
    body_html = ""

    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))

            # Skip attachments - we only want inline body content
            if "attachment" in content_disposition:
                continue

            if content_type == "text/plain" and not body_plain:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    try:
                        body_plain = payload.decode(charset, errors="replace")
                    except (LookupError, UnicodeDecodeError):
                        body_plain = payload.decode("utf-8", errors="replace")

            elif content_type == "text/html" and not body_html:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    try:
                        body_html = payload.decode(charset, errors="replace")
                    except (LookupError, UnicodeDecodeError):
                        body_html = payload.decode("utf-8", errors="replace")
    else:
        # Single-part message
        content_type = message.get_content_type()
        payload = message.get_payload(decode=True)

        if payload:
            charset = message.get_content_charset() or "utf-8"
            try:
                decoded_payload = payload.decode(charset, errors="replace")
            except (LookupError, UnicodeDecodeError):
                decoded_payload = payload.decode("utf-8", errors="replace")

            if content_type == "text/html":
                body_html = decoded_payload
            else:
                # Default to plain text for unknown content types
                body_plain = decoded_payload

    return body_plain, body_html


def parse_email_message(message: email.message.Message) -> EmailContent:
    """
    Parse an email.message.Message into an EmailContent dataclass.

    Consolidates all message parsing into a single function that extracts
    headers, addresses, and body content into the standard EmailContent
    structure.

    Args:
        message: Parsed email.message.Message object from mailbox iteration.

    Returns:
        EmailContent: Fully populated dataclass with all extracted data.

    Raises:
        TypeError: If parsedate_to_datetime receives None (missing Date header).
            This can happen with malformed emails. Consider adding validation.

    Implementation Notes:
        - Date is parsed via parsedate_to_datetime and converted to ISO format
        - All header values are decoded for RFC 2047 encoded words
        - Address headers (From, To, Cc, Bcc, Reply-To) are parsed specially
        - Body extraction delegates to get_body_content()

    Maintenance Considerations:
        - Missing Date header will raise TypeError - may need error handling
        - Message-ID may be missing in draft or malformed emails
        - In-Reply-To header links to parent message in thread
    """
    # Extract body content first
    body_plain, body_html = get_body_content(message)

    # Parse date - expects RFC 2822 format
    parsed_date = parsedate_to_datetime(
        decode_header_value(message.get("Date"))
    ).isoformat()

    # Build metadata with date and message ID
    metadata = EmailMetadata(
        date=parsed_date,
        message_id=decode_header_value(message.get("Message-ID", "")),
    )

    # Parse all address fields
    from_email = parse_email_address(message.get("From"))
    to_emails = parse_email_addresses(message.get("To"))
    to_cc = parse_email_addresses(message.get("Cc"))
    to_bcc = parse_email_addresses(message.get("Bcc"))
    reply_to = parse_email_addresses(message.get("Reply-To"))

    return EmailContent(
        from_email=from_email,
        subject=decode_header_value(message.get("Subject")),
        in_reply_to=decode_header_value(message.get("In-Reply-To", "")),
        reply_to=reply_to,
        to_emails=to_emails,
        to_cc=to_cc,
        to_bcc=to_bcc,
        body_plain=body_plain,
        body_html=body_html,
        metadata=metadata,
    )


def read_mbox_format_mail(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    """
    Read all emails from an mbox format file.

    Primary entry point for mbox extraction. Iterates through all messages
    in the mailbox and yields EmailContent objects for each.

    Args:
        file_like: BytesIO object containing complete mbox file data.
            The entire mbox content, potentially containing many messages.
        path: Optional filesystem path to source file. If provided, populates
            file metadata (filename, extension, folder) in each returned
            EmailContent.metadata. Useful for batch processing and auditing.

    Yields:
        EmailContent: One object per message in the mbox. Order matches
            the order of messages in the file.

    Raises:
        Various email parsing exceptions for malformed messages.

    Implementation Notes:
        This function parses mbox format entirely in memory:
        1. Splits the mbox data on "From " separator lines
        2. Parses each message with email.message_from_bytes()
        3. No temporary files are created (faster than disk I/O)

    Performance Considerations:
        - All parsing happens in memory (no disk I/O)
        - Memory usage is proportional to mbox file size
        - Messages are processed sequentially

    Example:
        >>> import io
        >>> # Mbox with two messages
        >>> mbox_content = b'''From sender@x.com Mon Jan 1 00:00:00 2024
        ... From: sender@x.com
        ... To: recipient@x.com
        ... Subject: First
        ...
        ... Body 1
        ...
        ... From other@x.com Tue Jan 2 00:00:00 2024
        ... From: other@x.com
        ... To: recipient@x.com
        ... Subject: Second
        ...
        ... Body 2
        ... '''
        >>> for email in read_mbox_format_mail(io.BytesIO(mbox_content)):
        ...     print(email.subject)
        First
        Second
    """
    try:
        file_like.seek(0)
        data = file_like.read()

        # Split mbox into individual messages in memory
        message_bytes_list = _split_mbox_messages(data)

        message_count = 0
        for msg_bytes in message_bytes_list:
            # Parse message from bytes using standard library
            message = email.message_from_bytes(msg_bytes)
            m = parse_email_message(message)

            if path:
                m.metadata.populate_from_path(path)

            message_count += 1
            logger.debug(
                "Extracted message %d: subject=%s",
                message_count,
                m.subject[:50] if m.subject else "(no subject)",
            )
            yield m

        logger.info("Extracted MBOX: %d messages", message_count)
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract MBOX file", cause=exc) from exc
