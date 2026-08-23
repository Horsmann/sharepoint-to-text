"""Construct canonical documents from parsed email values."""

from __future__ import annotations

from sharepoint2text.parsing.extractors._model import source_metadata
from sharepoint2text.parsing.models import (
    Attachment,
    ContentUnit,
    DocumentMetadata,
    ExtractedDocument,
    JsonValue,
)

Address = tuple[str, str]


def address_value(address: Address) -> dict[str, JsonValue]:
    """Convert an internal address pair to JSON-compatible properties.

    Args:
        address: Display-name and mailbox pair.

    Returns:
        Address fields suitable for a namespaced document property.
    """
    return {"name": address[0], "address": address[1]}


def build_email_document(
    *,
    source_format: str,
    path: str | None,
    sender: Address,
    subject: str,
    body_plain: str,
    body_html: str,
    date: str,
    message_id: str,
    in_reply_to: str,
    reply_to: list[Address],
    recipients: list[Address],
    cc: list[Address],
    bcc: list[Address],
    attachments: list[Attachment],
) -> ExtractedDocument:
    """Build one normalized email document from parsed header and body values.

    Args:
        source_format: Lowercase email container format.
        path: Optional source path.
        sender: Sender display-name and mailbox pair.
        subject: Message subject.
        body_plain: Plain-text body.
        body_html: HTML body fallback.
        date: Source-provided message timestamp.
        message_id: Message identifier header.
        in_reply_to: Parent-message identifier.
        reply_to: Reply-to recipients.
        recipients: Primary recipients.
        cc: Carbon-copy recipients.
        bcc: Blind-carbon-copy recipients.
        attachments: Canonical attachment records.

    Returns:
        Canonical message document.
    """
    properties: dict[str, JsonValue] = {
        f"{source_format}.from_email": address_value(sender),
    }
    scalar_values = {"in_reply_to": in_reply_to}
    for name, value in scalar_values.items():
        if value:
            properties[f"{source_format}.{name}"] = value
    for name, values in (
        ("reply_to", reply_to),
        ("to_emails", recipients),
        ("to_cc", cc),
        ("to_bcc", bcc),
    ):
        if values:
            properties[f"{source_format}.{name}"] = [
                address_value(value) for value in values
            ]

    if body_plain:
        properties[f"{source_format}.body_plain"] = body_plain.strip()
    if body_html:
        properties[f"{source_format}.body_html"] = body_html

    body = body_plain.strip() or body_html
    body_type = "plain" if body_plain.strip() else "html" if body_html else "empty"
    author = sender[1] or sender[0] or None
    return ExtractedDocument(
        format=source_format,
        source=source_metadata(path),
        metadata=DocumentMetadata(
            title=subject.strip() or None,
            author=author,
            created=date or None,
            properties=(
                {f"{source_format}.message_id": message_id} if message_id else {}
            ),
        ),
        units=[
            ContentUnit(
                number=1,
                kind="message",
                text=body,
                properties={f"{source_format}.body_type": body_type},
            )
        ],
        attachments=attachments,
        properties=properties,
    )
