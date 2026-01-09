"""
Native Python 7z Archive Support.

Pure Python implementation of 7z archive reading without external dependencies.
Uses Python's built-in lzma module for LZMA/LZMA2 decompression.

Based on the official 7-Zip SDK documentation.
"""

import io
import lzma
import os
import struct
import zlib
from dataclasses import dataclass
from typing import BinaryIO, Dict, List, Optional, Tuple

# 7z Magic Signature
MAGIC = b"7z\xbc\xaf\x27\x1c"

# Property IDs for 7z header parsing
PROP_END = 0x00
PROP_HEADER = 0x01
PROP_ARCHIVE_PROPERTIES = 0x02
PROP_ADDITIONAL_STREAMS_INFO = 0x03
PROP_MAIN_STREAMS_INFO = 0x04
PROP_FILES_INFO = 0x05
PROP_PACK_INFO = 0x06
PROP_UNPACK_INFO = 0x07
PROP_SUBSTREAMS_INFO = 0x08
PROP_SIZE = 0x09
PROP_CRC = 0x0A
PROP_FOLDER = 0x0B
PROP_CODERS_UNPACK_SIZE = 0x0C
PROP_NUM_UNPACK_STREAM = 0x0D
PROP_EMPTY_STREAM = 0x0E
PROP_EMPTY_FILE = 0x0F
PROP_NAME = 0x11
PROP_WIN_ATTRIBUTES = 0x15
PROP_ENCODED_HEADER = 0x17

# Coder IDs
CODER_COPY = b"\x00"
CODER_LZMA = b"\x03\x01\x01"
CODER_LZMA2 = b"\x21"
CODER_BCJ = b"\x03\x03\x01\x03"
CODER_AES_PREFIX = b"\x06\xf1\x07"


class Bad7zFile(Exception):
    """Exception raised for invalid or unsupported 7z files."""

    pass


@dataclass
class FileInfo:
    """Metadata for a file entry in the archive."""

    filename: str
    uncompressed: int
    is_directory: bool
    crc: Optional[int] = None
    attributes: int = 0
    folder_index: int = 0


@dataclass
class Folder:
    """A compression unit containing one or more files."""

    coders: List[Tuple[bytes, Optional[bytes]]]  # (coder_id, properties)
    unpack_sizes: List[int]
    crc: Optional[int] = None
    num_streams: int = 1


class SevenZipReader:
    """
    Low-level reader for 7z archives using only Python standard library.

    Supports:
    - LZMA compression
    - LZMA2 compression
    - Copy (uncompressed) method

    Not supported:
    - Encrypted archives (AES)
    - Split archives
    """

    def __init__(self, file: BinaryIO):
        self.file = file
        self._original_file = file
        self._files: List[FileInfo] = []
        self._folders: List[Folder] = []
        self._pack_positions: List[int] = []
        self._pack_sizes: List[int] = []
        self._file_sizes: List[int] = []
        self._header_offset = 0
        self._folder_to_files: Dict[int, List[int]] = {}

        self._parse_header()

    # -------------------------------------------------------------------------
    # Binary reading helpers
    # -------------------------------------------------------------------------

    def _read_bytes(self, n: int) -> bytes:
        """Read exactly n bytes, raising on EOF."""
        data = self.file.read(n)
        if len(data) != n:
            raise Bad7zFile(
                f"Unexpected end of file (expected {n} bytes, got {len(data)})"
            )
        return data

    def _read_uint8(self) -> int:
        return struct.unpack("<B", self._read_bytes(1))[0]

    def _read_uint32(self) -> int:
        return struct.unpack("<I", self._read_bytes(4))[0]

    def _read_uint64(self) -> int:
        return struct.unpack("<Q", self._read_bytes(8))[0]

    def _read_number(self) -> int:
        """Read a 7z variable-length encoded number."""
        first_byte = self._read_uint8()
        mask = 0x80
        value = 0

        for i in range(8):
            if (first_byte & mask) == 0:
                value |= (first_byte & (mask - 1)) << (i * 8)
                return value
            value |= self._read_uint8() << (i * 8)
            mask >>= 1

        return value

    def _read_boolean_vector(
        self, count: int, check_defined: bool = False
    ) -> List[bool]:
        """Read a packed boolean vector."""
        if check_defined:
            all_defined = self._read_uint8()
            if all_defined != 0:
                return [True] * count

        result = []
        byte_value = 0
        mask = 0

        for _ in range(count):
            if mask == 0:
                byte_value = self._read_uint8()
                mask = 0x80
            result.append((byte_value & mask) != 0)
            mask >>= 1

        return result

    def _seek_back_one(self) -> None:
        """Move file position back by one byte."""
        self.file.seek(self.file.tell() - 1)

    # -------------------------------------------------------------------------
    # Header parsing
    # -------------------------------------------------------------------------

    def _parse_header(self) -> None:
        """Parse the 7z archive signature and locate the end header."""
        self.file.seek(0)

        # Verify magic signature
        if self._read_bytes(6) != MAGIC:
            raise Bad7zFile("Invalid 7z signature")

        # Verify version (must be 0.x where x <= 4)
        major, minor = self._read_uint8(), self._read_uint8()
        if major != 0 or minor > 4:
            raise Bad7zFile(f"Unsupported 7z version: {major}.{minor}")

        # Read start header fields
        start_header_crc = self._read_uint32()
        next_header_offset = self._read_uint64()
        next_header_size = self._read_uint64()
        next_header_crc = self._read_uint32()

        # Verify start header CRC (covers 20 bytes after the CRC field)
        self.file.seek(12)
        if zlib.crc32(self.file.read(20)) & 0xFFFFFFFF != start_header_crc:
            raise Bad7zFile("Start header CRC mismatch")

        # Read and verify end header
        self._header_offset = 32
        header_pos = self._header_offset + next_header_offset

        self.file.seek(header_pos)
        header_data = self.file.read(next_header_size)

        if len(header_data) != next_header_size:
            raise Bad7zFile("Could not read full header")

        if zlib.crc32(header_data) & 0xFFFFFFFF != next_header_crc:
            raise Bad7zFile("Header CRC mismatch")

        # Parse the header content
        self.file = io.BytesIO(header_data)
        self._parse_end_header()

    def _parse_end_header(self) -> None:
        """Parse the end header, handling compressed headers if needed."""
        prop_id = self._read_uint8()

        if prop_id == PROP_ENCODED_HEADER:
            self._parse_encoded_header()
            prop_id = self._read_uint8()

        if prop_id == PROP_HEADER:
            self._parse_main_header()
        elif prop_id != PROP_END:
            raise Bad7zFile(f"Unexpected property ID: {prop_id}")

    def _parse_encoded_header(self) -> None:
        """Decompress a compressed header."""
        pack_info = self._parse_pack_info()
        unpack_info = self._parse_unpack_info()

        if not unpack_info:
            raise Bad7zFile("No unpack info in encoded header")

        # Skip optional substreams info
        prop_id = self._read_uint8()
        if prop_id == PROP_SUBSTREAMS_INFO:
            self._skip_substreams_info(unpack_info)
            prop_id = self._read_uint8()

        if prop_id != PROP_END:
            raise Bad7zFile(f"Expected END property, got {prop_id}")

        # Decompress the header
        pack_pos, pack_sizes = pack_info or (self._header_offset, [])

        saved_file = self.file
        self.file = self._original_file
        try:
            decompressed = self._decompress_folder(unpack_info[0], pack_pos, pack_sizes)
        finally:
            self.file = saved_file

        self.file = io.BytesIO(decompressed)

    def _parse_main_header(self) -> None:
        """Parse the main header containing streams and files info."""
        prop_id = self._read_uint8()

        # Skip archive properties
        if prop_id == PROP_ARCHIVE_PROPERTIES:
            while True:
                prop_id = self._read_uint8()
                if prop_id == PROP_END:
                    break
                size = self._read_number()
                self._read_bytes(size)
            prop_id = self._read_uint8()

        # Skip additional streams info
        if prop_id == PROP_ADDITIONAL_STREAMS_INFO:
            self._parse_streams_info()
            prop_id = self._read_uint8()

        # Parse main streams info
        if prop_id == PROP_MAIN_STREAMS_INFO:
            self._parse_streams_info()
            prop_id = self._read_uint8()

        # Parse files info
        if prop_id == PROP_FILES_INFO:
            self._parse_files_info()
            prop_id = self._read_uint8()

        if prop_id != PROP_END:
            raise Bad7zFile(f"Expected END, got {prop_id}")

    # -------------------------------------------------------------------------
    # Streams info parsing
    # -------------------------------------------------------------------------

    def _parse_streams_info(self) -> None:
        """Parse the streams info section."""
        prop_id = self._read_uint8()

        if prop_id == PROP_PACK_INFO:
            self._seek_back_one()
            self._parse_pack_info()
            prop_id = self._read_uint8()

        if prop_id == PROP_UNPACK_INFO:
            self._seek_back_one()
            self._parse_unpack_info()
            prop_id = self._read_uint8()

        if prop_id == PROP_SUBSTREAMS_INFO:
            self._parse_substreams_info()
            prop_id = self._read_uint8()

        if prop_id != PROP_END:
            raise Bad7zFile(f"Expected END in streams info, got {prop_id}")

    def _parse_pack_info(self) -> Optional[Tuple[int, List[int]]]:
        """Parse pack stream information (compressed data locations)."""
        prop_id = self._read_uint8()
        if prop_id != PROP_PACK_INFO:
            self._seek_back_one()
            return None

        pack_pos = self._read_number()
        num_pack_streams = self._read_number()
        pack_sizes: List[int] = []

        prop_id = self._read_uint8()

        if prop_id == PROP_SIZE:
            pack_sizes = [self._read_number() for _ in range(num_pack_streams)]
            prop_id = self._read_uint8()

        if prop_id == PROP_CRC:
            defined = self._read_boolean_vector(num_pack_streams, check_defined=True)
            for is_defined in defined:
                if is_defined:
                    self._read_uint32()
            prop_id = self._read_uint8()

        if prop_id != PROP_END:
            raise Bad7zFile(f"Expected END in pack info, got {prop_id}")

        absolute_pos = pack_pos + self._header_offset
        self._pack_positions = [absolute_pos]
        self._pack_sizes = pack_sizes

        return (absolute_pos, pack_sizes)

    def _parse_unpack_info(self) -> List[Folder]:
        """Parse unpack information (decompression settings)."""
        prop_id = self._read_uint8()
        if prop_id != PROP_UNPACK_INFO:
            self._seek_back_one()
            return []

        prop_id = self._read_uint8()
        if prop_id != PROP_FOLDER:
            raise Bad7zFile("Expected FOLDER property in unpack info")

        num_folders = self._read_number()

        # External indicator (not supported)
        if self._read_uint8() != 0:
            raise Bad7zFile("External folders not supported")

        folders = [self._parse_folder() for _ in range(num_folders)]

        # Read unpack sizes for each coder
        prop_id = self._read_uint8()
        if prop_id != PROP_CODERS_UNPACK_SIZE:
            raise Bad7zFile("Expected CODERS_UNPACK_SIZE")

        for folder in folders:
            folder.unpack_sizes = [
                self._read_number() for _ in range(len(folder.coders))
            ]

        # Read optional CRCs
        prop_id = self._read_uint8()
        if prop_id == PROP_CRC:
            defined = self._read_boolean_vector(num_folders, check_defined=True)
            for i, is_defined in enumerate(defined):
                if is_defined:
                    folders[i].crc = self._read_uint32()
            prop_id = self._read_uint8()

        if prop_id != PROP_END:
            raise Bad7zFile(f"Expected END in unpack info, got {prop_id}")

        self._folders = folders
        return folders

    def _parse_folder(self) -> Folder:
        """Parse a single folder (compression unit)."""
        num_coders = self._read_number()
        coders: List[Tuple[bytes, Optional[bytes]]] = []

        for _ in range(num_coders):
            flags = self._read_uint8()
            coder_id_size = flags & 0x0F
            is_complex = (flags & 0x10) != 0
            has_attributes = (flags & 0x20) != 0

            coder_id = self._read_bytes(coder_id_size)

            # Skip stream counts for complex coders
            if is_complex:
                self._read_number()  # num_in_streams
                self._read_number()  # num_out_streams

            properties = None
            if has_attributes:
                props_size = self._read_number()
                properties = self._read_bytes(props_size)

            coders.append((coder_id, properties))

        # Skip bind pairs
        num_bind_pairs = len(coders) - 1
        for _ in range(num_bind_pairs):
            self._read_number()  # in_index
            self._read_number()  # out_index

        return Folder(coders=coders, unpack_sizes=[])

    def _parse_substreams_info(self) -> None:
        """Parse substreams info for multi-file folders."""
        prop_id = self._read_uint8()
        file_sizes: List[int] = []

        # Number of streams per folder
        if prop_id == PROP_NUM_UNPACK_STREAM:
            for folder in self._folders:
                folder.num_streams = self._read_number()
            prop_id = self._read_uint8()
        else:
            for folder in self._folders:
                folder.num_streams = 1

        # Individual file sizes within folders
        if prop_id == PROP_SIZE:
            for folder in self._folders:
                total = folder.unpack_sizes[-1] if folder.unpack_sizes else 0
                for _ in range(folder.num_streams - 1):
                    size = self._read_number()
                    file_sizes.append(size)
                    total -= size
                if total > 0:
                    file_sizes.append(total)
            prop_id = self._read_uint8()
        else:
            for folder in self._folders:
                if folder.unpack_sizes:
                    file_sizes.append(folder.unpack_sizes[-1])

        self._file_sizes = file_sizes

        # Skip CRCs
        if prop_id == PROP_CRC:
            total_streams = sum(f.num_streams for f in self._folders)
            defined = self._read_boolean_vector(total_streams, check_defined=True)
            for is_defined in defined:
                if is_defined:
                    self._read_uint32()
            prop_id = self._read_uint8()

        if prop_id != PROP_END:
            raise Bad7zFile(f"Expected END in substreams info, got {prop_id}")

    def _skip_substreams_info(self, folders: List[Folder]) -> None:
        """Skip substreams info during encoded header parsing."""
        prop_id = self._read_uint8()

        if prop_id == PROP_NUM_UNPACK_STREAM:
            for folder in folders:
                folder.num_streams = self._read_number()
            prop_id = self._read_uint8()

        if prop_id == PROP_SIZE:
            for folder in folders:
                for _ in range(folder.num_streams - 1):
                    self._read_number()
            prop_id = self._read_uint8()

        if prop_id == PROP_CRC:
            total_streams = sum(f.num_streams for f in folders)
            defined = self._read_boolean_vector(total_streams, check_defined=True)
            for is_defined in defined:
                if is_defined:
                    self._read_uint32()
            prop_id = self._read_uint8()

        if prop_id != PROP_END:
            raise Bad7zFile(f"Expected END in substreams info, got {prop_id}")

    # -------------------------------------------------------------------------
    # Files info parsing
    # -------------------------------------------------------------------------

    def _parse_files_info(self) -> None:
        """Parse the files information section."""
        num_files = self._read_number()
        empty_streams: List[bool] = [False] * num_files
        names: List[str] = [""] * num_files
        attributes: List[int] = [0] * num_files

        while True:
            prop_id = self._read_uint8()
            if prop_id == PROP_END:
                break

            size = self._read_number()
            end_pos = self.file.tell() + size

            if prop_id == PROP_EMPTY_STREAM:
                empty_streams = self._read_boolean_vector(num_files)

            elif prop_id == PROP_EMPTY_FILE:
                # Skip empty file markers (not needed for extraction)
                pass

            elif prop_id == PROP_NAME:
                if self._read_uint8() != 0:
                    raise Bad7zFile("External names not supported")
                for i in range(num_files):
                    name_chars = []
                    while True:
                        char = struct.unpack("<H", self._read_bytes(2))[0]
                        if char == 0:
                            break
                        name_chars.append(chr(char))
                    names[i] = "".join(name_chars)

            elif prop_id == PROP_WIN_ATTRIBUTES:
                defined = self._read_boolean_vector(num_files, check_defined=True)
                for i, is_defined in enumerate(defined):
                    if is_defined:
                        attributes[i] = self._read_uint32()

            # Ensure correct position for next property
            self.file.seek(end_pos)

        self._build_file_list(num_files, empty_streams, names, attributes)

    def _build_file_list(
        self,
        num_files: int,
        empty_streams: List[bool],
        names: List[str],
        attributes: List[int],
    ) -> None:
        """Build the file list from parsed metadata."""
        size_index = 0

        for i in range(num_files):
            is_dir = empty_streams[i] or (attributes[i] & 0x10) != 0

            if is_dir:
                file_size = 0
            elif size_index < len(self._file_sizes):
                file_size = self._file_sizes[size_index]
                size_index += 1
            else:
                file_size = 0

            self._files.append(
                FileInfo(
                    filename=names[i],
                    uncompressed=file_size,
                    is_directory=is_dir,
                    attributes=attributes[i],
                )
            )

        # Map non-directory files to folders
        folder_idx = 0
        file_in_folder = 0

        for i, file_info in enumerate(self._files):
            if file_info.is_directory or folder_idx >= len(self._folders):
                continue

            file_info.folder_index = folder_idx
            self._folder_to_files.setdefault(folder_idx, []).append(i)

            file_in_folder += 1
            if file_in_folder >= self._folders[folder_idx].num_streams:
                folder_idx += 1
                file_in_folder = 0

    # -------------------------------------------------------------------------
    # Decompression
    # -------------------------------------------------------------------------

    def _decompress_folder(
        self,
        folder: Folder,
        pack_pos: int,
        pack_sizes: List[int],
        source_file: Optional[BinaryIO] = None,
    ) -> bytes:
        """Decompress all data in a folder."""
        if not folder.coders:
            raise Bad7zFile("No coders in folder")

        file_to_read = source_file or self.file

        # Read compressed data
        total_size = sum(pack_sizes) if pack_sizes else 0
        if total_size == 0:
            file_to_read.seek(0, 2)
            total_size = file_to_read.tell() - pack_pos

        file_to_read.seek(pack_pos)
        data = file_to_read.read(total_size)

        # Apply decoders in reverse order
        for coder_id, properties in reversed(folder.coders):
            data = self._apply_decoder(coder_id, properties, data, folder.unpack_sizes)

        return data

    def _apply_decoder(
        self,
        coder_id: bytes,
        properties: Optional[bytes],
        data: bytes,
        unpack_sizes: List[int],
    ) -> bytes:
        """Apply a single decoder to data."""
        if coder_id == CODER_COPY:
            return data

        if coder_id == CODER_LZMA:
            return self._decompress_lzma(data, properties, unpack_sizes)

        if coder_id == CODER_LZMA2:
            return self._decompress_lzma2(data, properties)

        if coder_id == CODER_BCJ:
            return data  # BCJ filter pass-through for text extraction

        if coder_id.startswith(CODER_AES_PREFIX):
            raise Bad7zFile("Encrypted archives are not supported")

        raise Bad7zFile(f"Unsupported compression method: {coder_id.hex()}")

    def _decompress_lzma(
        self, data: bytes, properties: Optional[bytes], unpack_sizes: List[int]
    ) -> bytes:
        """Decompress LZMA data."""
        if properties is None or len(properties) < 5:
            raise Bad7zFile("Invalid LZMA properties")

        unpack_size = unpack_sizes[-1] if unpack_sizes else -1
        size_bytes = struct.pack("<Q", unpack_size) if unpack_size >= 0 else b"\xff" * 8

        # Construct LZMA alone format: props (5 bytes) + size (8 bytes) + data
        lzma_stream = properties[:5] + size_bytes + data

        try:
            return lzma.LZMADecompressor(format=lzma.FORMAT_ALONE).decompress(
                lzma_stream
            )
        except lzma.LZMAError as e:
            raise Bad7zFile(f"LZMA decompression failed: {e}") from e

    def _decompress_lzma2(self, data: bytes, properties: Optional[bytes]) -> bytes:
        """Decompress LZMA2 data."""
        if properties is None or len(properties) < 1:
            raise Bad7zFile("LZMA2 requires properties")

        prop_byte = properties[0]

        # Calculate dictionary size from properties byte
        if prop_byte < 40:
            if prop_byte > 0:
                dict_size = (2 | (prop_byte & 1)) << (prop_byte // 2 + 11)
            else:
                dict_size = 1 << 12
            filters = [{"id": lzma.FILTER_LZMA2, "dict_size": dict_size}]
        else:
            filters = [{"id": lzma.FILTER_LZMA2, "preset": 6}]

        try:
            return lzma.LZMADecompressor(
                format=lzma.FORMAT_RAW, filters=filters
            ).decompress(data)
        except lzma.LZMAError as e:
            raise Bad7zFile(f"LZMA2 decompression failed: {e}") from e

    # -------------------------------------------------------------------------
    # Public API
    # -------------------------------------------------------------------------

    def list(self) -> List[FileInfo]:
        """Return a copy of the file list."""
        return self._files.copy()

    def needs_password(self) -> bool:
        """Check if the archive uses AES encryption."""
        return any(
            coder_id.startswith(CODER_AES_PREFIX)
            for folder in self._folders
            for coder_id, _ in folder.coders
        )


class SevenZipFile:
    """
    Context manager for reading 7z archives.

    Provides a py7zr-compatible interface for drop-in replacement.

    Example:
        with SevenZipFile(file_like, "r") as szf:
            for info in szf.list():
                print(info.filename)
            szf.extractall(path=temp_dir)
    """

    def __init__(self, file: BinaryIO, mode: str = "r", password: Optional[str] = None):
        if mode != "r":
            raise Bad7zFile(f"Mode '{mode}' not supported, only 'r' is available")

        self._file = file
        self._reader: Optional[SevenZipReader] = None

    def __enter__(self) -> "SevenZipFile":
        self._reader = SevenZipReader(self._file)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self._reader = None

    def list(self) -> List[FileInfo]:
        """List all files in the archive."""
        if self._reader is None:
            raise Bad7zFile("Archive not opened")
        return self._reader.list()

    def extractall(self, path: str) -> None:
        """Extract all files to the specified directory."""
        if self._reader is None:
            raise Bad7zFile("Archive not opened")

        os.makedirs(path, exist_ok=True)

        for folder_idx, folder in enumerate(self._reader._folders):
            if folder_idx not in self._reader._folder_to_files:
                continue

            pack_pos = (
                self._reader._pack_positions[0]
                if self._reader._pack_positions
                else self._reader._header_offset
            )

            try:
                decompressed = self._reader._decompress_folder(
                    folder,
                    pack_pos,
                    self._reader._pack_sizes,
                    source_file=self._file,
                )
            except Bad7zFile as e:
                raise Bad7zFile(f"Failed to decompress folder {folder_idx}: {e}") from e

            self._extract_files_from_folder(path, folder_idx, decompressed)

    def _extract_files_from_folder(
        self, base_path: str, folder_idx: int, decompressed: bytes
    ) -> None:
        """Extract individual files from decompressed folder data."""
        offset = 0

        for file_idx in self._reader._folder_to_files[folder_idx]:
            file_info = self._reader._files[file_idx]

            if file_info.is_directory:
                os.makedirs(os.path.join(base_path, file_info.filename), exist_ok=True)
                continue

            file_path = os.path.join(base_path, file_info.filename)

            parent_dir = os.path.dirname(file_path)
            if parent_dir:
                os.makedirs(parent_dir, exist_ok=True)

            file_data = decompressed[offset : offset + file_info.uncompressed]
            offset += file_info.uncompressed

            with open(file_path, "wb") as f:
                f.write(file_data)

    def needs_password(self) -> bool:
        """Check if the archive requires a password."""
        if self._reader is None:
            raise Bad7zFile("Archive not opened")
        return self._reader.needs_password()
