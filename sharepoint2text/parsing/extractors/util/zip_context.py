import io

from sharepoint2text.parsing import _defused_xml as ET
from sharepoint2text.parsing.extractors.util.zip_bomb import open_zipfile
from sharepoint2text.parsing.extractors.util.zip_utils import (
    read_zip_text,
    read_zip_xml_root,
)

XmlElement = ET.Element


class ZipContext:
    """Reusable ZIP context with convenience helpers for reading OOXML/ODF files."""

    def __init__(self, file_like: io.BytesIO):
        self.file_like = file_like
        self.file_like.seek(0)
        self._zip = open_zipfile(self.file_like, source=type(self).__name__)
        self._namelist = set(self._zip.namelist())

    @property
    def namelist(self) -> set[str]:
        """Return the set of member names available in the package.

        Returns:
            Set of package member names.
        """
        return self._namelist

    def exists(self, path: str) -> bool:
        """Return whether the requested member exists in the package.

        Args:
            path: Package-relative member path.

        Returns:
            True when the named package member is available.
        """
        return path in self._namelist

    def read_xml_root(self, path: str) -> XmlElement:
        """Parse and return the root element of an XML package member.

        Args:
            path: Package-relative XML member path.

        Returns:
            Safely parsed XML root element.
        """
        return read_zip_xml_root(self._zip, path)

    def read_text(self, path: str) -> str:
        """Decode and return a UTF-8 text package member.

        Args:
            path: Package-relative text member path.

        Returns:
            Decoded package text.
        """
        return read_zip_text(self._zip, path)

    def read_bytes(self, path: str) -> bytes:
        """Return the raw bytes stored for a package member.

        Args:
            path: Package-relative binary member path.

        Returns:
            Raw package member bytes.
        """
        return self._zip.read(path)

    def open_stream(self, path: str):  # type: ignore[no-untyped-def]
        """Open a readable stream for a package member.

        Args:
            path: Package-relative member path.

        Returns:
            Readable binary member stream owned by the package.
        """
        return self._zip.open(path)

    def close(self) -> None:
        """Close the underlying package and release its resources.

        Returns:
            None.
        """
        self._zip.close()
