"""
utils/format_detection.py
─────────────────────────
Detect file format from a URL or raw bytes (magic bytes).
"""

import os
from urllib.parse import urlparse, unquote


_MAGIC: list[tuple[bytes, str]] = [
    (b"%PDF", "pdf"),
    (b"PK\x03\x04", "docx_or_xlsx"),
    (b"\xd0\xcf\x11\xe0", "doc_or_xls"),
    (b"\xff\xd8\xff", "jpg"),
    (b"\x89PNG\r\n\x1a\n", "png"),
    (b"GIF87a", "gif"),
    (b"GIF89a", "gif"),
    (b"BM", "bmp"),
    (b"II\x2a\x00", "tiff"),
    (b"MM\x00\x2a", "tiff"),
]

_EXTENSION_MAP: dict[str, str] = {
    ".pdf": "pdf",
    ".docx": "docx",
    ".doc": "doc",
    ".xlsx": "xlsx",
    ".xls": "xls",
    ".jpg": "jpg",
    ".jpeg": "jpg",
    ".png": "png",
    ".gif": "gif",
    ".bmp": "bmp",
    ".tiff": "tiff",
    ".tif": "tiff",
    ".txt": "txt",
    ".log": "txt",
    ".csv": "csv",
}

SUPPORTED_FORMATS = set(_EXTENSION_MAP.values())


def detect_format_from_url(url: str) -> str:
    """
    Derive format from the file extension in the URL path.
    Returns a lower-case format string or 'unknown'.
    """
    path = unquote(urlparse(url).path)
    _, ext = os.path.splitext(path)
    return _EXTENSION_MAP.get(ext.lower(), "unknown")


def detect_format_from_bytes(data: bytes) -> str:
    """
    Derive format from the file's magic bytes.
    Falls back to 'unknown'.  For ZIP-based Office formats (docx/xlsx)
    we peek at the internal structure to distinguish them.
    """
    for magic, fmt in _MAGIC:
        if data[:len(magic)] == magic:
            if fmt == "docx_or_xlsx":
                return _distinguish_zip_office(data)
            if fmt == "doc_or_xls":
                return _distinguish_ole_office(data)
            return fmt

    # Heuristic: printable ASCII → plain text
    try:
        sample = data[:512].decode("utf-8", errors="strict")
        if sample.isprintable() or "\n" in sample:
            return "txt"
    except UnicodeDecodeError:
        pass

    return "unknown"


def _distinguish_zip_office(data: bytes) -> str:
    """Peek inside the ZIP to tell .docx from .xlsx."""
    try:
        import zipfile, io
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = zf.namelist()
        if any(n.startswith("word/") for n in names):
            return "docx"
        if any(n.startswith("xl/") for n in names):
            return "xlsx"
    except Exception:
        pass
    return "docx"  # Safe default


def _distinguish_ole_office(data: bytes) -> str:
    """Very rough heuristic for OLE2 files."""
    try:
        # Excel OLE files contain the string 'Workbook' or 'Book'
        if b"W\x00o\x00r\x00k\x00b\x00o\x00o\x00k" in data[:4096]:
            return "xls"
    except Exception:
        pass
    return "doc"


def detect_format(url: str, data: bytes) -> str:
    """
    Try URL extension first (fast); fall back to magic bytes.
    """
    fmt = detect_format_from_url(url)
    if fmt != "unknown":
        return fmt
    return detect_format_from_bytes(data)
