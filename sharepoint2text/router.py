import io
import logging
import mimetypes
import os
from typing import Any, Callable, Generator

from sharepoint2text.exceptions import ExtractionFileFormatNotSupportedError
from sharepoint2text.extractors.data_types import ExtractionInterface

logger = logging.getLogger(__name__)

mime_type_mapping = {
    # legacy ms
    "application/vnd.ms-powerpoint": "ppt",
    "application/vnd.ms-excel": "xls",
    "application/msword": "doc",
    "application/rtf": "rtf",
    "text/rtf": "rtf",
    # modern ms
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    # open office
    "application/vnd.oasis.opendocument.text": "odt",
    "application/vnd.oasis.opendocument.presentation": "odp",
    "application/vnd.oasis.opendocument.spreadsheet": "ods",
    # plain text
    "text/csv": "csv",
    "application/csv": "csv",
    "application/json": "json",
    "text/json": "json",
    "text/plain": "txt",
    "text/markdown": "md",
    "text/tab-separated-values": "tsv",
    "application/tab-separated-values": "tsv",
    # email
    "application/vnd.ms-outlook": "msg",
    "message/rfc822": "eml",
    "application/mbox": "mbox",
    # other
    "text/html": "html",
    "application/xhtml+xml": "html",
    "application/pdf": "pdf",
}


def _get_extractor(
    file_type: str,
) -> Callable[[io.BytesIO, str | None], Generator[ExtractionInterface, Any, None]]:
    """Return the extractor function for a file type (lazy import)."""
    if file_type == "xlsx":
        from sharepoint2text.extractors.ms_modern.xlsx_extractor import read_xlsx

        return read_xlsx
    elif file_type == "xls":
        from sharepoint2text.extractors.ms_legacy.xls_extractor import read_xls

        return read_xls
    elif file_type == "ppt":
        from sharepoint2text.extractors.ms_legacy.ppt_extractor import read_ppt

        return read_ppt
    elif file_type == "pptx":
        from sharepoint2text.extractors.ms_modern.pptx_extractor import read_pptx

        return read_pptx
    elif file_type == "doc":
        from sharepoint2text.extractors.ms_legacy.doc_extractor import read_doc

        return read_doc
    elif file_type == "docx":
        from sharepoint2text.extractors.ms_modern.docx_extractor import read_docx

        return read_docx
    elif file_type == "pdf":
        from sharepoint2text.extractors.pdf_extractor import read_pdf

        return read_pdf
    elif file_type in ("csv", "json", "txt", "tsv", "md"):
        from sharepoint2text.extractors.plain_extractor import read_plain_text

        return read_plain_text
    elif file_type == "msg":
        from sharepoint2text.extractors.mail.msg_email_extractor import (
            read_msg_format_mail,
        )

        return read_msg_format_mail
    elif file_type == "mbox":
        from sharepoint2text.extractors.mail.mbox_email_extractor import (
            read_mbox_format_mail,
        )

        return read_mbox_format_mail
    elif file_type == "eml":
        from sharepoint2text.extractors.mail.eml_email_extractor import (
            read_eml_format_mail,
        )

        return read_eml_format_mail
    elif file_type == "rtf":
        from sharepoint2text.extractors.ms_legacy.rtf_extractor import read_rtf

        return read_rtf
    elif file_type == "html":
        from sharepoint2text.extractors.html_extractor import read_html

        return read_html
    elif file_type == "odt":
        from sharepoint2text.extractors.open_office.odt_extractor import read_odt

        return read_odt
    elif file_type == "odp":
        from sharepoint2text.extractors.open_office.odp_extractor import read_odp

        return read_odp
    elif file_type == "ods":
        from sharepoint2text.extractors.open_office.ods_extractor import read_ods

        return read_ods
    else:
        raise ExtractionFileFormatNotSupportedError(
            f"No extractor for file type: {file_type}"
        )


def is_supported_file(path: str) -> bool:
    """Checks if the path is a supported file"""
    path = path.lower()
    mime_type, _ = mimetypes.guess_type(path)
    return mime_type in mime_type_mapping or any(
        [path.endswith(ending) for ending in [".msg", ".eml", ".mbox", ".md"]]
    )


def get_extractor(
    path: str,
) -> Callable[[io.BytesIO, str | None], Generator[ExtractionInterface, Any, None]]:
    """Analysis the path of a file and returns a suited extractor.
       The file MUST not exist (yet). The path or filename alone suffices to return an
       extractor.

    :returns a function of an extractor. All extractors take a file-like object as parameter
    :raises ExtractionFileFormatNotSupportedError: File is not covered by any extractor
    """
    path = path.lower()

    mime_type, _ = mimetypes.guess_type(path)
    logger.debug(f"Guessed mime type: [{mime_type}]")
    if mime_type is not None and mime_type in mime_type_mapping:
        file_type = mime_type_mapping[mime_type]
        logger.debug(
            f"Detected file type: {file_type} (MIME: {mime_type}) for file: {path}"
        )
        return _get_extractor(file_type)
    elif any([path.endswith(ending) for ending in [".msg", ".eml", ".mbox", ".md"]]):
        # the file types are mapped with leading dot
        path_elements = os.path.splitext(path)
        if len(path_elements) <= 1:
            raise ExtractionFileFormatNotSupportedError(
                f"The file path did not allow to identify the file type [{path}]"
            )
        file_type = path_elements[1][1:]
        logger.debug(f"Detected file type: {file_type} for file: {path}")
        return _get_extractor(file_type)
    else:
        logger.debug(f"File [{path}] with mime type [{mime_type}] is not supported")
        raise ExtractionFileFormatNotSupportedError(
            f"File type not supported: {mime_type}"
        )
