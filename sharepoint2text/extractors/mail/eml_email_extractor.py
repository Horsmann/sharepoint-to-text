import io
import logging
from typing import Any, Generator

from mailparser import parse_from_bytes

from sharepoint2text.extractors.data_types import (
    EmailAddress,
    EmailContent,
    EmailMetadata,
)

logger = logging.getLogger(__name__)


def _read_eml_format(payload: bytes) -> EmailContent:
    """Parse .eml file bytes and return EmailContent."""
    mail = parse_from_bytes(payload)

    # Extract addresses
    from_email = EmailAddress(mail.from_[0][0], mail.from_[0][1])
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

    # Extract date as string
    date_str = ""
    if mail.date:
        date_str = mail.date.isoformat()

    metadata = EmailMetadata(
        date=date_str,
        message_id=mail.message_id or "",
    )

    # Body extraction - mailparser uses text_plain for plain text
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
        metadata=metadata,
    )


def read_eml_format_mail(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    """Read an .eml file and extract its content.

    Args:
        file_like: BytesIO object containing the email data
        path: Optional path to populate file metadata

    Returns:
        Generator of EmailContent objects. This accounts for some email formats containing multiple emails.
    """
    file_like.seek(0)
    content = _read_eml_format(file_like.getvalue())

    if path:
        content.metadata.populate_from_path(path)

    yield content
