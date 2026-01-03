"""
Mail Extractor Package
======================

This package provides extractors for parsing and extracting text content from
various email file formats commonly encountered in document management systems
and legacy archives.

Supported Formats
-----------------

.eml (RFC 5322 / MIME):
    Standard Internet Message Format emails. Uses the `mailparser` library
    for robust parsing of MIME structures.

.mbox (Unix Mailbox):
    Unix mailbox format containing multiple concatenated emails. Uses Python's
    built-in `mailbox` module. Requires temporary file creation for processing.

.msg (Microsoft Outlook):
    Proprietary Microsoft Outlook message format based on OLE compound documents.
    Uses the `msg_parser` library (MsOxMessage) for extraction.

Common Output Structure
-----------------------
All extractors return `EmailContent` objects (via generator) containing:
    - from_email: Sender EmailAddress (name, address)
    - to_emails: List of recipient EmailAddress objects
    - to_cc, to_bcc: Carbon copy and blind carbon copy recipients
    - reply_to: Reply-to addresses
    - subject: Email subject line
    - body_plain: Plain text body content
    - body_html: HTML body content (if available)
    - metadata: EmailMetadata with date, message_id, and file info

Usage Example
-------------
    >>> from sharepoint2text.parsing.extractors.mail import read_eml_format_mail
    >>> import io
    >>>
    >>> with open("message.eml", "rb") as f:
    ...     data = io.BytesIO(f.read())
    ...     for email in read_eml_format_mail(data, path="message.eml"):
    ...         print(email.subject)

Dependencies
------------
- mailparser: EML parsing (pip install mail-parser)
- msg_parser: MSG parsing (pip install msg_parser)
- Python stdlib: mailbox, email modules for mbox

Maintenance Notes
-----------------
- All extractors follow the generator pattern for consistency, even for
  single-message formats (EML, MSG). This allows uniform iteration.
- Character encoding is handled with fallback to UTF-8 with replacement
  characters for malformed data.
- Date parsing expects RFC 2822 format; non-standard dates may cause issues.

See Also
--------
- sharepoint2text.parsing.extractors.data_types: EmailContent, EmailAddress, EmailMetadata
"""

from sharepoint2text.parsing.extractors.mail.eml_email_extractor import (
    read_eml_format_mail,
)
from sharepoint2text.parsing.extractors.mail.mbox_email_extractor import (
    read_mbox_format_mail,
)
from sharepoint2text.parsing.extractors.mail.msg_email_extractor import (
    read_msg_format_mail,
)

__all__ = [
    "read_eml_format_mail",
    "read_mbox_format_mail",
    "read_msg_format_mail",
]
