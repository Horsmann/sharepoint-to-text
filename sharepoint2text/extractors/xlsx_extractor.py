"""
XLSX content extractor using pandas and openpyxl libraries.
"""

import datetime
import io
import logging
from typing import Any, Generator, List

import pandas as pd
from openpyxl import load_workbook

from sharepoint2text.extractors.data_types import XlsxContent, XlsxMetadata, XlsxSheet

logger = logging.getLogger(__name__)


def _read_metadata(file_like: io.BytesIO) -> XlsxMetadata:
    file_like.seek(0)
    wb = load_workbook(file_like)
    props = wb.properties

    metadata = XlsxMetadata(
        title=props.title or "",
        description=props.description or "",
        creator=props.creator or "",
        last_modified_by=props.lastModifiedBy or "",
        created=(
            props.created.isoformat()
            if isinstance(props.created, datetime.datetime)
            else ""
        ),
        modified=(
            props.modified.isoformat()
            if isinstance(props.modified, datetime.datetime)
            else ""
        ),
        keywords=props.keywords or "",
        language=props.language or "",
        revision=props.revision,
    )
    wb.close()
    return metadata


def _read_content(file_like: io.BytesIO) -> List[XlsxSheet]:
    logger.debug("Reading content")
    file_like.seek(0)
    xls = pd.read_excel(file_like, engine="calamine", sheet_name=None)

    sheets = []
    for sheet_name, df in xls.items():
        logger.debug(f"Reading sheet: [{sheet_name}]")
        data = df.to_dict(orient="records")
        text = df.to_string(index=False)
        sheets.append(
            XlsxSheet(
                name=str(sheet_name),
                data=data,
                text=text,
            )
        )
    return sheets


def read_xlsx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[XlsxContent, Any, None]:
    """
    Extract all relevant content from an XLSX file.

    Args:
        file_like: A BytesIO object containing the XLSX file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        MicrosoftXlsxContent dataclass with all extracted content.
    """
    sheets = _read_content(file_like)
    metadata = _read_metadata(file_like)
    metadata.populate_from_path(path)

    yield XlsxContent(metadata=metadata, sheets=sheets)
