import io
import logging
import re
from email.utils import parsedate_to_datetime
from typing import Any, Generator

from msg_parser import MsOxMessage

from sharepoint2text.extractors.data_types import (
    EmailAddress,
    EmailContent,
    EmailMetadata,
)

logger = logging.getLogger(__name__)


def _parse_single_recipient(raw: str) -> EmailAddress | None:
    """
    Parse a recipient string in the format 'Name <email@example.com>'.

    Examples:
        'Benny Bottema <benny@bennybottema.com>' -> EmailAddress(name='Benny Bottema', address='benny@bennybottema.com')
        '<benny@bennybottema.com>' -> EmailAddress(name='', address='benny@bennybottema.com')
        'benny@bennybottema.com' -> EmailAddress(name='', address='benny@bennybottema.com')
        'Benny Bottema' -> EmailAddress(name='Benny Bottema', address='')
    """
    raw = raw.strip()
    if not raw:
        return None

    # Look for <email> at the end of the string
    match = re.search(r"<([^>]+)>\s*$", raw)
    if match:
        address = match.group(1).strip()
        name = raw[: match.start()].strip().strip("\"'")
        return EmailAddress(name=name, address=address)

    # No angle brackets - check if it's just an email
    if "@" in raw and " " not in raw:
        return EmailAddress(name="", address=raw)

    # No email found, treat as name only
    return EmailAddress(name=raw, address="")


def _parse_multi_recipients(raw: str | list[str]) -> list[EmailAddress]:
    """
    Parse recipient string(s) that may contain multiple addresses.

    Handles both:
        - List input: ['Benny Bottema <benny@bennybottema.com>']
        - String input: 'Alice <alice@example.com>; Bob <bob@example.com>'
    """
    if not raw:
        return []

    # Handle list input
    if isinstance(raw, list):
        results = []
        for item in raw:
            results.extend(_parse_multi_recipients(item))
        return results

    # Split by semicolon or comma (common separators)
    parts = re.split(r"[;,]", raw)

    recipients = []
    for part in parts:
        addr = _parse_single_recipient(part)
        if addr and (addr.name or addr.address):
            recipients.append(addr)

    return recipients


def read_msg_format_mail(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    msg = MsOxMessage(file_like)

    meta = EmailMetadata(
        message_id=msg.message_id,
        date=parsedate_to_datetime(msg.sent_date).isoformat(),
    )

    content = EmailContent(
        subject=msg.subject,
        from_email=_parse_multi_recipients(msg.sender)[0],
        to_emails=_parse_multi_recipients(msg.to),
        to_cc=_parse_multi_recipients(msg.cc),
        to_bcc=_parse_multi_recipients(msg.bcc),
        reply_to=msg.reply_to,
        body_plain=msg.body,
        body_html=msg.body,
        metadata=meta,
    )

    if path:
        content.metadata.populate_from_path(path)

    yield content
