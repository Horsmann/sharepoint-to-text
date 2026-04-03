import io
import tarfile
import zipfile


def read_file_to_file_like(path: str) -> io.BytesIO:
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)
        return file_like


def _zip_bytes_to_file_like(files: dict[str, str]) -> io.BytesIO:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, text in files.items():
            zf.writestr(name, text)
    buffer.seek(0)
    return buffer


def tar_bytes_to_file_like(files: dict[str, bytes]) -> io.BytesIO:
    buffer = io.BytesIO()
    with tarfile.open(fileobj=buffer, mode="w") as tf:
        for name, data in files.items():
            info = tarfile.TarInfo(name=name)
            info.size = len(data)
            tf.addfile(info, io.BytesIO(data))
    buffer.seek(0)
    return buffer
