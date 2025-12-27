import email
import email.header
import email.utils
import io
import mailbox
import os
import tempfile
from email.utils import parsedate_to_datetime
from typing import Any, Generator

from sharepoint2text.extractors.data_types import (
    EmailAddress,
    EmailContent,
    EmailMetadata,
)


def decode_header_value(value: str | None) -> str:
    """Decode an email header value, handling encoded words."""
    if not value:
        return ""

    decoded_parts = []
    for part, charset in email.header.decode_header(value):
        if isinstance(part, bytes):
            charset = charset or "utf-8"
            try:
                decoded_parts.append(part.decode(charset, errors="replace"))
            except (LookupError, UnicodeDecodeError):
                decoded_parts.append(part.decode("utf-8", errors="replace"))
        else:
            decoded_parts.append(part)

    return "".join(decoded_parts)


def parse_email_address(addr_string: str | None) -> EmailAddress:
    """Parse a single email address string into an EmailAddress object."""
    if not addr_string:
        return EmailAddress()

    name, address = email.utils.parseaddr(addr_string)
    return EmailAddress(name=decode_header_value(name), address=address)


def parse_email_addresses(addr_string: str | None) -> list[EmailAddress]:
    """Parse a comma-separated list of email addresses."""
    if not addr_string:
        return []

    addresses = email.utils.getaddresses([addr_string])
    result = []
    for name, addr in addresses:
        if addr:  # Only include if there's an actual address
            result.append(EmailAddress(name=decode_header_value(name), address=addr))
    return result


def get_body_content(message: email.message.Message) -> tuple[str, str]:
    """Extract plain text and HTML body from an email message."""
    body_plain = ""
    body_html = ""

    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))

            # Skip attachments
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
                body_plain = decoded_payload

    return body_plain, body_html


def parse_email_message(message: email.message.Message) -> EmailContent:
    """Parse an email.message.Message into an EmailContent dataclass."""

    # Extract body content
    body_plain, body_html = get_body_content(message)

    parsed_date = parsedate_to_datetime(
        decode_header_value(message.get("Date"))
    ).isoformat()

    # Parse metadata
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

    Args:
        file_like: A BytesIO object containing mbox data
        path: Optional path hint (not used for parsing)

    Returns:
        List of EmailContent objects for each email in the mbox
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".mbox") as tmp:
        tmp.write(file_like.read())
        tmp_path = tmp.name

    try:
        mbox = mailbox.mbox(tmp_path)
        for message in mbox:
            m = parse_email_message(message)
            if path:
                m.metadata.populate_from_path(path)
            yield m

    finally:
        os.unlink(tmp_path)
