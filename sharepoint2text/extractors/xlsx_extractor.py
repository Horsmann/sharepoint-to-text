import datetime
import io
import logging

import pandas as pd
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


def _read_metadata(file_like: io.BytesIO) -> dict:
    file_like.seek(0)
    wb = load_workbook(file_like)
    props = wb.properties

    metadata = {
        "title": props.title,
        "description": props.description,
        "lastModifiedBy": props.lastModifiedBy,
        "keywords": props.keywords,
        "language": props.language,
        "revision": props.revision,
        "creator": props.creator,
        "created": (
            props.created.isoformat()
            if isinstance(props.created, datetime.datetime)
            else None
        ),
        "modified": (
            props.modified.isoformat()
            if isinstance(props.modified, datetime.datetime)
            else None
        ),
    }
    wb.close()
    return metadata


def _read_content(file_like: io.BytesIO):
    logger.debug("Reading content")
    file_like.seek(0)
    xls = pd.read_excel(file_like, engine="calamine", sheet_name=None)

    data = {}
    for sheet_name, df in xls.items():
        logger.debug(f"Reading sheet: [{sheet_name}]")
        data[sheet_name] = df.to_dict(orient="records")
    return data


def read_xlsx(file_like: io.BytesIO) -> dict:
    content = _read_content(file_like)
    metadata = _read_metadata(file_like)

    return {"metadata": metadata, "content": content}
