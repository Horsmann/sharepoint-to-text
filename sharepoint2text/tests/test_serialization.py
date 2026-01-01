import io
import json
import logging
import unittest

from sharepoint2text import read_file
from sharepoint2text.extractors.data_types import PptxContent

logger = logging.getLogger(__name__)

tc = unittest.TestCase()


def _read_file_to_file_like(path: str) -> io.BytesIO:
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)
        return file_like


def test_serialize_for_json() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"
    obj = next(read_file(path))
    assert isinstance(obj, PptxContent)

    payload = obj.to_json()
    assert isinstance(payload, dict)

    try:
        json.dumps(payload)
    except Exception as e:
        tc.fail("Unexpected exception: {}".format(e))
