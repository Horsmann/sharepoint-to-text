"""
XLS content extractor using pandas and olefile libraries.
"""

import io
import logging
from typing import Any, Generator, List

import olefile
import pandas as pd

from sharepoint2text.extractors.data_types import XlsContent, XlsMetadata, XlsSheet

logger = logging.getLogger(__name__)


def _read_content(file_like: io.BytesIO) -> List[XlsSheet]:
    logger.debug("Reading content")
    xls = pd.read_excel(file_like, engine="calamine", sheet_name=None)

    sheets = []
    for sheet_name, df in xls.items():
        logger.debug(f"Reading sheet: [{sheet_name}]")
        data = df.to_dict(orient="records")
        text = df.to_string(index=False)
        sheets.append(
            XlsSheet(
                name=str(sheet_name),
                data=data,
                text=text,
            )
        )
    return sheets


def _read_metadata(file_like: io.BytesIO) -> XlsMetadata:
    ole = olefile.OleFileIO(file_like)
    meta = ole.get_metadata()

    result = XlsMetadata(
        title=meta.title.decode("utf-8") if meta.title else "",
        author=meta.author.decode("utf-8") if meta.author else "",
        subject=meta.subject.decode("utf-8") if meta.subject else "",
        company=meta.company.decode("utf-8") if meta.company else "",
        last_saved_by=meta.last_saved_by.decode("utf-8") if meta.last_saved_by else "",
        created=meta.create_time.isoformat() if meta.create_time else "",
        modified=meta.last_saved_time.isoformat() if meta.last_saved_time else "",
    )
    ole.close()
    return result


def read_xls(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[XlsContent, Any, None]:
    """
    Extract all relevant content from an XLS file.

    Args:
        file_like: A BytesIO object containing the XLS file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        MicrosoftXlsContent dataclass with all extracted content.
    """
    file_like.seek(0)
    sheets = _read_content(file_like=file_like)
    file_like.seek(0)
    metadata = _read_metadata(file_like=file_like)
    metadata.populate_from_path(path)

    full_text = "\n\n".join(sheet.text for sheet in sheets)

    yield XlsContent(
        metadata=metadata,
        sheets=sheets,
        full_text=full_text,
    )
