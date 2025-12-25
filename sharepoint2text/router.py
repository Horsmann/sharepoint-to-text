import logging
import mimetypes
import typing

from sharepoint2text.extractors.doc_extractor import read_doc
from sharepoint2text.extractors.docx_extractor import read_docx
from sharepoint2text.extractors.pdf_extractor import read_pdf
from sharepoint2text.extractors.plain_extractor import read_plain_text
from sharepoint2text.extractors.ppt_extractor import read_ppt
from sharepoint2text.extractors.pptx_extractor import read_pptx
from sharepoint2text.extractors.xls_extractor import read_xls
from sharepoint2text.extractors.xlsx_extractor import read_xlsx

logger = logging.getLogger(__name__)

mime_type_mapping = {
    "application/vnd.ms-powerpoint": "ppt",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
    "application/msword": "doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
    "application/vnd.ms-excel": "xls",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "application/pdf": "pdf",
    "text/csv": "csv",
    "application/csv": "csv",
    "application/json": "json",
    "text/json": "json",
    "text/plain": "txt",
    "text/tab-separated-values": "tsv",
    "application/tab-separated-values": "tsv",
}

extractor_mappings = {
    "xlsx": read_xlsx,
    "xls": read_xls,
    "ppt": read_ppt,
    "pptx": read_pptx,
    "doc": read_doc,
    "docx": read_docx,
    "pdf": read_pdf,
    "csv": read_plain_text,
    "json": read_plain_text,
    "txt": read_plain_text,
    "tsv": read_plain_text,
}


def get_extractor(path: str) -> typing.Callable:
    """Analysis the path of a file and returns a suited extractor.
       The file MUST not exist (yet). The path or filename alone suffices to return an
       extractor.

    :returns a function of an extractor. All extractors take a file-like object as parameter
    :raises RuntimeError: File is not covered by any extractor
    """

    mime_type, _ = mimetypes.guess_type(path)

    if mime_type in mime_type_mapping:
        file_type = mime_type_mapping[mime_type]
        logger.debug(
            f"Detected file type: {file_type} (MIME: {mime_type}) for file: {path}"
        )
        return extractor_mappings[file_type]

    else:
        logger.debug(f"File [{path}] with mime type [{mime_type}] is not supported")
        raise RuntimeError(f"File type not supported: {mime_type}")
