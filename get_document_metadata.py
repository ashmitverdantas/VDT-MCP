"""──────────────────────────────
MCP Tool: getDocumentMetadata

Returns format, size, and content hash of a document WITHOUT
extracting its full text. Useful for a quick file inspection
before committing to a full analyzeDocuments call.
"""

from __future__ import annotations
import hashlib
import logging
from datetime import datetime, timezone
from typing import Any
from mcp.types import Tool

from sharepoint import download_file
from format_detection import detect_format

logger = logging.getLogger(__name__)

GET_METADATA_TOOL = Tool(
    name="getDocumentMetadata",
    description=(
        "Fetch lightweight metadata for a SharePoint document: format, file size, "
        "SHA-256 hash, and whether text extraction is supported. "
        "Does NOT return document text — use analyzeDocuments for full extraction."
    ),
    inputSchema={
        "type": "object",
        "required": ["documentUrl"],
        "properties": {
            "documentUrl": {
                "type": "string",
                "description": "SharePoint URL of the document.",
            }
        },
    },
)

SUPPORTED_FORMATS = {
    "pdf", "docx", "doc", "xlsx", "xls",
    "jpg", "png", "gif", "bmp", "tiff", "txt", "csv",
}


async def handle_get_document_metadata(args: dict) -> dict[str, Any]:
    '''Handler for getDocumentMetadata tool calls. Returns metadata or error info.'''
    url: str = args.get("documentUrl", "").strip()
    if not url:
        return {"error": "documentUrl is required."}

    try:
        data = await download_file(url)
        fmt = detect_format(url, data)
        return {
            "schema_version": "1.0",
            "tool": "getDocumentMetadata",
            "inspected_at": datetime.now(timezone.utc).isoformat(),
            "url": url,
            "format": fmt,
            "file_size_bytes": len(data),
            "content_hash_sha256": hashlib.sha256(data).hexdigest(),
            "extraction_supported": fmt in SUPPORTED_FORMATS,
        }
    except (ValueError, OSError, RuntimeError) as exc:
        logger.error("getDocumentMetadata failed for %s: %s", url, exc, exc_info=True)
        return {
            "schema_version": "1.0",
            "tool": "getDocumentMetadata",
            "inspected_at": datetime.now(timezone.utc).isoformat(),
            "url": url,
            "error": str(exc),
        }
