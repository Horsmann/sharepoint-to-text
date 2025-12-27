"""
Legacy .doc Reader (Word 97-2003 / OLE2 Binary Format)

Usage:
    from doc_reader import read_doc, MicrosoftDocContent

    with MicrosoftDocContent('file.doc') as doc:
        text = doc.read()
        main = doc.get_main_text()
        hf = doc.get_headers_footers()
        fn = doc.get_footnotes()
        meta = doc.get_metadata()

DISCLAIMER:
    the module is 100% AI-generated with Claude Opus 4.5
    output is blackbox tested but implementation details
    may be subject to improvements
"""

import datetime
import io
import logging
import re
import struct
from typing import Any, Generator, List, Optional

import olefile

from sharepoint2text.extractors.data_types import (
    DocContent,
    DocMetadata,
)

logger = logging.getLogger(__name__)


def read_doc(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[DocContent, Any, None]:
    """
    Extract all relevant content from a DOC file.

    Args:
        file_like: A BytesIO object containing the DOC file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        MicrosoftDocContent dataclass with all extracted content.
    """
    file_like.seek(0)
    with _DocReader(file_like) as doc:
        document = doc.read()
        document.metadata = doc.get_metadata()
        document.metadata.populate_from_path(path)
        yield document


class _DocReader:
    def __init__(self, file_like: io.BytesIO):
        self.file_like = file_like
        self.ole = None
        self._content: Optional[DocContent] = None
        self._is_unicode: Optional[bool] = None
        self._text_start: Optional[int] = None

    def __enter__(self):
        self.ole = olefile.OleFileIO(self.file_like)
        return self

    def __exit__(self, *args):
        if self.ole:
            self.ole.close()

    def _get_stream(self, name: str) -> bytes:
        if self.ole and self.ole.exists(name):
            return self.ole.openstream(name).read()
        return b""

    def _parse_content(self) -> DocContent:
        if self._content is not None:
            return self._content

        if not self.ole:
            raise RuntimeError("File not opened")

        word_doc = self._get_stream("WordDocument")
        if not word_doc:
            raise ValueError("No WordDocument Stream")

        if len(word_doc) < 0x200:
            raise ValueError("File too small")

        # Magic check
        magic = struct.unpack_from("<H", word_doc, 0)[0]
        if magic != 0xA5EC:
            raise ValueError(f"Not a valid.doc file (Magic: {hex(magic)})")

        # Check flags
        flags = struct.unpack_from("<H", word_doc, 0x0A)[0]
        if flags & 0x0100:
            raise ValueError("Fils is encrypted")

        # Character counts aus FIB
        # Main text
        ccp_text = struct.unpack_from("<I", word_doc, 0x4C)[0]
        # Footnotes
        ccp_ftn = struct.unpack_from("<I", word_doc, 0x50)[0]
        # Headers/Footers
        ccp_hdd = struct.unpack_from("<I", word_doc, 0x54)[0]
        # Annotations
        ccp_atn = struct.unpack_from("<I", word_doc, 0x5C)[0]

        self._text_start, self._is_unicode = self._find_text_start_and_enc(word_doc)

        # Byte-multiplicator (2 for UTF-16LE, 1 for CP1252)
        mult = 2 if self._is_unicode else 1
        encoding = "utf-16-le" if self._is_unicode else "cp1252"

        pos = self._text_start

        # Main text
        main_data = word_doc[pos : pos + ccp_text * mult]
        pos += ccp_text * mult

        # Footnotes
        ftn_data = word_doc[pos : pos + ccp_ftn * mult] if ccp_ftn > 0 else b""
        pos += ccp_ftn * mult

        # Headers/Footers
        hdd_data = word_doc[pos : pos + ccp_hdd * mult] if ccp_hdd > 0 else b""
        pos += ccp_hdd * mult

        # Annotations
        atn_data = word_doc[pos : pos + ccp_atn * mult] if ccp_atn > 0 else b""

        self._content = DocContent(
            main_text=self._clean_text(main_data.decode(encoding, errors="replace")),
            footnotes=(
                self._clean_text(ftn_data.decode(encoding, errors="replace"))
                if ftn_data
                else ""
            ),
            headers_footers=(
                self._clean_text(hdd_data.decode(encoding, errors="replace"))
                if hdd_data
                else ""
            ),
            annotations=(
                self._clean_text(atn_data.decode(encoding, errors="replace"))
                if atn_data
                else ""
            ),
        )

        return self._content

    def read(self) -> DocContent:
        """
        Extracts the text from the document.
        """
        content = self._parse_content()

        return content

    def get_main_text(self) -> str:
        return self._parse_content().main_text

    def get_headers_footers(self) -> str:
        return self._parse_content().headers_footers

    def get_footnotes(self) -> str:
        return self._parse_content().footnotes

    def get_annotations(self) -> str:
        return self._parse_content().annotations

    def get_all_parts(self) -> DocContent:
        return self._parse_content()

    @staticmethod
    def _find_text_start_and_enc(word_doc: bytes) -> tuple:
        """
        Finds the start of the text and detects whether it is UTF-16LE or CP1252.

        Returns:
            (offset, is_unicode): Start position and True if UTF-16LE
        """
        for offset in range(0x200, min(len(word_doc) - 64, 0x2000), 0x40):
            sample = word_doc[offset : offset + 64]

            utf16_score = 0
            for i in range(0, min(len(sample) - 1, 60), 2):
                b1, b2 = sample[i], sample[i + 1]
                if (0x20 <= b1 <= 0x7E or b1 in (0x0D, 0x0A)) and b2 == 0x00:
                    utf16_score += 1
                elif b1 in (0xE4, 0xF6, 0xFC, 0xC4, 0xD6, 0xDC, 0xDF) and b2 == 0x00:
                    utf16_score += 1

            cp1252_score = sum(
                1
                for b in sample
                if (0x20 <= b <= 0x7E) or b in (0x0D, 0x0A, 0x09) or (0xC0 <= b <= 0xFF)
            )

            if utf16_score > 20:
                return offset, True
            elif cp1252_score > 45:
                return offset, False

        return 0x800, False

    @staticmethod
    def _clean_text(text: str) -> str:
        if not text:
            return ""

        replacements = {
            "\x07": "\t",
            "\x0b": "\n",
            "\x0c": "\n\n",
            "\x0d": "\n",
            "\x13": "",
            "\x14": " ",
            "\x15": "",
            "\x01": "",
            "\x08": "",
            "\x19": "",
            "\x1e": "",
            "\x1f": "",
            "\xa0": " ",
        }
        for old, new in replacements.items():
            text = text.replace(old, new)

        text = re.sub(r"[\x00-\x08\x0e-\x1f\x7f]", "", text)
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    def get_metadata(self) -> DocMetadata:
        if not self.ole:
            return DocMetadata()
        try:
            m = self.ole.get_metadata()
            return DocMetadata(
                title=m.title.decode("utf-8"),
                author=m.author.decode("utf-8"),
                subject=m.subject.decode("utf-8"),
                keywords=m.keywords.decode("utf-8"),
                last_saved_by=m.last_saved_by.decode("utf-8"),
                create_time=(
                    m.create_time.isoformat()
                    if isinstance(m.create_time, datetime.datetime)
                    else ""
                ),
                last_saved_time=(
                    m.last_saved_time.isoformat()
                    if isinstance(m.last_saved_time, datetime.datetime)
                    else ""
                ),
                num_pages=m.num_pages,
                num_words=m.num_words,
                num_chars=m.num_chars,
            )
        except Exception as e:
            logger.debug(f"Metadata extraction failed: [{e}]")
            return DocMetadata()

    def list_streams(self) -> List[List[str]]:
        return self.ole.listdir() if self.ole else []
