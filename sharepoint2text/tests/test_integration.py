import glob
import logging
import os
import unittest

from sharepoint2text import read_file

logger = logging.getLogger(__name__)

tc = unittest.TestCase()


def test_read_file():
    for path in glob.glob("sharepoint2text/tests/resources/**/*", recursive=True):
        if not os.path.isfile(path):
            continue
        logger.debug(f"Calling read_file with: [{path}]")
        for obj in read_file(path=path):
            # verify that all obj have the ExtractionInterface methods
            tc.assertTrue(hasattr(obj, "get_metadata"))
            tc.assertTrue(hasattr(obj, "iterator"))
            tc.assertTrue(hasattr(obj, "get_full_text"))
