import io
import zipfile

import olefile

from sharepoint2text.extractors.util.zip_bomb import open_zipfile


def _has_ole_encryption_stream(ole: olefile.OleFileIO) -> bool:
    for stream in ("EncryptionInfo", "EncryptedPackage", "DataSpaces"):
        if ole.exists(stream):
            return True
    return False


def is_ooxml_encrypted(file_like: io.BytesIO) -> bool:
    file_like.seek(0)
    if olefile.isOleFile(file_like):
        file_like.seek(0)
        with olefile.OleFileIO(file_like) as ole:
            encrypted = _has_ole_encryption_stream(ole)
        file_like.seek(0)
        return encrypted
    file_like.seek(0)
    return False


def is_odf_encrypted(file_like: io.BytesIO) -> bool:
    file_like.seek(0)
    if not zipfile.is_zipfile(file_like):
        file_like.seek(0)
        return False

    file_like.seek(0)
    with open_zipfile(file_like, source="is_odf_encrypted") as zf:
        try:
            manifest = zf.read("META-INF/manifest.xml").decode("utf-8", errors="ignore")
        except KeyError:
            return False

    encrypted = (
        "encryption-data" in manifest
        or "manifest:encrypted" in manifest
        or "manifest:algorithm" in manifest
    )
    file_like.seek(0)
    return encrypted


def is_xls_encrypted(file_like: io.BytesIO) -> bool:
    file_like.seek(0)
    if not olefile.isOleFile(file_like):
        file_like.seek(0)
        return False

    file_like.seek(0)
    with olefile.OleFileIO(file_like) as ole:
        stream_name = None
        if ole.exists("Workbook"):
            stream_name = "Workbook"
        elif ole.exists("Book"):
            stream_name = "Book"

        if not stream_name:
            file_like.seek(0)
            return False

        data = ole.openstream(stream_name).read()

    offset = 0
    data_len = len(data)
    while offset + 4 <= data_len:
        record_id = int.from_bytes(data[offset : offset + 2], "little")
        record_len = int.from_bytes(data[offset + 2 : offset + 4], "little")
        if record_id == 0x002F:  # FILEPASS
            file_like.seek(0)
            return True
        offset += 4 + record_len
    file_like.seek(0)
    return False


def is_ppt_encrypted(file_like: io.BytesIO) -> bool:
    file_like.seek(0)
    if not olefile.isOleFile(file_like):
        file_like.seek(0)
        return False

    file_like.seek(0)
    with olefile.OleFileIO(file_like) as ole:
        if _has_ole_encryption_stream(ole):
            file_like.seek(0)
            return True
        encrypted = ole.exists("EncryptedSummary") or ole.exists(
            "EncryptedSummaryInformation"
        )
    file_like.seek(0)
    return encrypted
