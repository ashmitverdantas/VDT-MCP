from __future__ import annotations
import asyncio
import hashlib
import logging
import time
from datetime import datetime, timezone
from typing import Any

from mcp.types import Tool

from sharepoint import download_file
from extractors import extract_text_from_bytes
from format_detection import detect_format

logger = logging.getLogger(__name__)

MAX_FILES = 10
MAX_FILE_BYTES = 50 * 1024 * 1024  # 50 MB


# ── Tool schema (MCP descriptor) ──────────────────────────────────────────────

ANALYZE_DOCUMENTS_TOOL = Tool(
    name="analyzeDocuments",
    description=(
        "Download one or more documents from SharePoint and extract their full text content. "
        "Supports PDF, Word (.docx/.doc), Excel (.xlsx/.xls), images (JPG/PNG, OCR), "
        "plain text, and CSV files. Returns structured JSON with extracted text and metadata. "
        "Use this to feed document content into an AI analysis pipeline or to surface "
        "EHS incident data stored in SharePoint."
    ),
    inputSchema={
        "type": "object",
        "required": ["documentUrls"],
        "properties": {
            "documentUrls": {
                "type": "array",
                "description": "List of SharePoint file URLs to analyse (1–10).",
                "minItems": 1,
                "maxItems": MAX_FILES,
                "items": {
                    "type": "string",
                    "description": "Direct download URL or SharePoint web URL of the file.",
                },
            },
            "includeRawText": {
                "type": "boolean",
                "default": True,
                "description": "Include the full extracted text in the response.",
            },
            "summarize": {
                "type": "boolean",
                "default": False,
                "description": "If true, Claude will append a one-paragraph summary per document.",
            },
        },
    },
)


# ── Handler ───────────────────────────────────────────────────────────────────

async def handle_analyze_documents(args: dict) -> dict[str, Any]:
    """Handle the analyzeDocuments tool request.
    Downloads and extracts text from SharePoint documents, returning structured
    JSON with extracted content and metadata.
    """
    urls: list[str] = args.get("documentUrls", [])
    include_raw: bool = args.get("includeRawText", True)

    if not urls:
        return _error_response("No documentUrls provided.")
    if len(urls) > MAX_FILES:
        return _error_response(f"Too many files: {len(urls)}. Maximum is {MAX_FILES}.")

    start = time.perf_counter()
    tasks = [_process_single(url, include_raw) for url in urls]
    results = await asyncio.gather(*tasks, return_exceptions=False)

    elapsed = round(time.perf_counter() - start, 2)
    succeeded = sum(1 for r in results if r["status"] == "success")
    failed = len(results) - succeeded

    return {
        "schema_version": "1.0",
        "tool": "analyzeDocuments",
        "processed_at": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "total": len(urls),
            "succeeded": succeeded,
            "failed": failed,
            "elapsed_seconds": elapsed,
        },
        "documents": results,
    }


# ── Per-file processing ───────────────────────────────────────────────────────

async def _process_single(url: str, include_raw: bool) -> dict[str, Any]:
    doc: dict[str, Any] = {
        "url": url,
        "status": "error",
        "format": None,
        "file_size_bytes": None,
        "content_hash": None,
        "text": None,
        "text_length": None,
        "error": None,
    }

    try:
        # 1. Download
        data = await download_file(url)
        doc["file_size_bytes"] = len(data)

        if len(data) > MAX_FILE_BYTES:
            raise ValueError(
                f"File exceeds {MAX_FILE_BYTES // (1024*1024)} MB limit "
                f"({len(data) // (1024*1024)} MB)."
            )

        # 2. Detect format
        fmt = detect_format(url, data)
        doc["format"] = fmt
        doc["content_hash"] = hashlib.sha256(data).hexdigest()

        # 3. Extract text
        extracted = extract_text_from_bytes(fmt, data)

        doc["status"] = "success"
        doc["text_length"] = len(extracted)
        if include_raw:
            doc["text"] = extracted

    except (ValueError, OSError, RuntimeError) as exc:
        logger.error("Failed to process %s: %s", url, exc, exc_info=True)
        doc["error"] = str(exc)

    return doc


# ── Helpers ───────────────────────────────────────────────────────────────────

def _error_response(message: str) -> dict:
    return {
        "schema_version": "1.0",
        "tool": "analyzeDocuments",
        "processed_at": datetime.now(timezone.utc).isoformat(),
        "error": message,
        "documents": [],
    }
