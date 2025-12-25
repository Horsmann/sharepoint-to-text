import io

import olefile
import pandas as pd


def _read_content(file_like: io.BytesIO):
    xls = pd.read_excel(file_like, engine="calamine", sheet_name=None)

    data = {}
    for sheet_name, df in xls.items():
        print(f"Sheet: {sheet_name}")
        data[sheet_name] = df.to_dict(orient="records")
    return data


def _read_metadata(file_like: io.BytesIO) -> dict:
    ole = olefile.OleFileIO(file_like)
    meta = ole.get_metadata()

    result = {
        "author": meta.author.decode("utf-8") if meta.author else "",
        "last_saved_by": meta.last_saved_by.decode("utf-8")
        if meta.last_saved_by
        else "",
        "created": meta.create_time.isoformat(),
        "modified": meta.last_saved_time.isoformat(),
        "title": meta.title.decode("utf-8") if meta.title else "",
        "subject": meta.subject.decode("utf-8") if meta.subject else "",
        "company": meta.company.decode("utf-8") if meta.company else "",
    }
    ole.close()
    return result


def read_xls(file_like: io.BytesIO) -> dict:
    file_like.seek(0)
    data = _read_content(file_like=file_like)
    file_like.seek(0)

    meta = _read_metadata(file_like=file_like)

    return {
        "sheets": data,
        "metadata": meta,
    }
