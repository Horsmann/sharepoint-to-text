import io
from xml.etree.ElementTree import Element as XmlElement

from sharepoint2text.extractors.util.zip_bomb import open_zipfile
from sharepoint2text.extractors.util.zip_utils import read_zip_text, read_zip_xml_root


class ZipContext:
    """Reusable ZIP context with convenience helpers for reading OOXML/ODF files."""

    def __init__(self, file_like: io.BytesIO):
        self.file_like = file_like
        self.file_like.seek(0)
        self._zip = open_zipfile(self.file_like, source=type(self).__name__)
        self._namelist = set(self._zip.namelist())

    @property
    def namelist(self) -> set[str]:
        return self._namelist

    def exists(self, path: str) -> bool:
        return path in self._namelist

    def read_xml_root(self, path: str) -> XmlElement:
        return read_zip_xml_root(self._zip, path)

    def read_text(self, path: str) -> str:
        return read_zip_text(self._zip, path)

    def read_bytes(self, path: str) -> bytes:
        return self._zip.read(path)

    def open_stream(self, path: str) -> io.BufferedReader:
        return self._zip.open(path)

    def close(self) -> None:
        self._zip.close()
