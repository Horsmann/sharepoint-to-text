import typing
from abc import abstractmethod
from typing import Protocol


class ExtractionInterface(Protocol):
    @abstractmethod
    def iterator(self) -> typing.Iterator[str]:
        """
        Returns an iterator over the extracted text i.e., the main text body of a file.
        Additional text areas may be missing if they are not part of the main text body of the file.
        This greatly depends on the underlying data source.
        A PDF returns text per pages, PowerPoint files return slides as units.
        Excel files return sheets.
        Content of footnotes, headers or alike is not part of this iterator's return values.
        The legacy and modern Word documents have no per-page representation in the files, they return only a single unit which is the full text.
        """
        ...

    def get_full_text(self) -> str:
        """Full text of the slide deck as one single block of text"""
        ...

    # @abstractmethod
    # def get_metadata(self) -> dict:
    #     """ Returns the metadata of the extracted file """
    #     ...
