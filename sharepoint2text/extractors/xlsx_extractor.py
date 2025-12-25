"""
XLSX content extractor using pandas and openpyxl libraries.
"""

import datetime
import io
import logging
import typing
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook

from sharepoint2text.extractors.abstract_extractor import ExtractionInterface

logger = logging.getLogger(__name__)


@dataclass
class MicrosoftXlsxMetadata:
    title: str = ""
    description: str = ""
    creator: str = ""
    last_modified_by: str = ""
    created: str = ""
    modified: str = ""
    keywords: str = ""
    language: str = ""
    revision: Optional[str] = None


@dataclass
class MicrosoftXlsxSheet:
    name: str = ""
    data: List[Dict[str, Any]] = field(default_factory=list)
    text: str = ""


@dataclass
class MicrosoftXlsxContent(ExtractionInterface):
    metadata: MicrosoftXlsxMetadata = field(default_factory=MicrosoftXlsxMetadata)
    sheets: List[MicrosoftXlsxSheet] = field(default_factory=list)

    def iterator(self) -> typing.Iterator[str]:
        for sheet in self.sheets:
            yield sheet.name + "\n" + sheet.text.strip()

    def get_full_text(self) -> str:
        return "\n".join(list(self.iterator()))


def _read_metadata(file_like: io.BytesIO) -> MicrosoftXlsxMetadata:
    file_like.seek(0)
    wb = load_workbook(file_like)
    props = wb.properties

    metadata = MicrosoftXlsxMetadata(
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


def _read_content(file_like: io.BytesIO) -> List[MicrosoftXlsxSheet]:
    logger.debug("Reading content")
    file_like.seek(0)
    xls = pd.read_excel(file_like, engine="calamine", sheet_name=None)

    sheets = []
    for sheet_name, df in xls.items():
        logger.debug(f"Reading sheet: [{sheet_name}]")
        data = df.to_dict(orient="records")
        text = df.to_string(index=False)
        sheets.append(
            MicrosoftXlsxSheet(
                name=str(sheet_name),
                data=data,
                text=text,
            )
        )
    return sheets


def read_xlsx(file_like: io.BytesIO) -> MicrosoftXlsxContent:
    """
    Extract all relevant content from an XLSX file.

    Args:
        file_like: A BytesIO object containing the XLSX file data.

    Returns:
        MicrosoftXlsxContent dataclass with all extracted content.
    """
    sheets = _read_content(file_like)
    metadata = _read_metadata(file_like)

    return MicrosoftXlsxContent(metadata=metadata, sheets=sheets)
