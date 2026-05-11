"""
tools/list_sharepoint_files.py
──────────────────────────────
MCP Tool: listSharePointFiles

Lists files in a SharePoint document library folder and returns
their metadata as structured JSON. Useful for agents that need to
discover which files exist before calling analyzeDocuments.
"""

from __future__ import annotations
import logging
from datetime import datetime, timezone
from typing import Any

from mcp.types import Tool

from sharepoint import list_files_in_folder

logger = logging.getLogger(__name__)

LIST_FILES_TOOL = Tool(
    name="listSharePointFiles",
    description=(
        "List files in a SharePoint document library. Returns file names, sizes, "
        "last-modified dates, and direct download URLs. Useful for discovering which "
        "EHS documents are available before calling analyzeDocuments."
    ),
    inputSchema={
        "type": "object",
        "required": ["siteId", "driveId"],
        "properties": {
            "siteId": {
                "type": "string",
                "description": "Microsoft Graph site ID (e.g. contoso.sharepoint.com,abc123,xyz456).",
            },
            "driveId": {
                "type": "string",
                "description": "Document library drive ID from Microsoft Graph.",
            },
            "folderPath": {
                "type": "string",
                "default": "root",
                "description": "Folder path relative to the drive root. Use 'root' for the library root.",
            },
            "fileTypes": {
                "type": "array",
                "items": {"type": "string"},
                "description": "Optional list of extensions to filter (e.g. [\".pdf\", \".docx\"]).",
            },
        },
    },
)


async def handle_list_sharepoint_files(args: dict) -> dict[str, Any]:
    site_id: str = args.get("siteId", "")
    drive_id: str = args.get("driveId", "")
    folder_path: str = args.get("folderPath", "root")
    file_types: list[str] | None = args.get("fileTypes") or None

    if not site_id or not drive_id:
        return {"error": "siteId and driveId are required.", "files": []}

    try:
        files = await list_files_in_folder(site_id, drive_id, folder_path, file_types)
        return {
            "schema_version": "1.0",
            "tool": "listSharePointFiles",
            "listed_at": datetime.now(timezone.utc).isoformat(),
            "folder_path": folder_path,
            "total_files": len(files),
            "files": files,
        }
    except Exception as exc:
        logger.error(f"listSharePointFiles failed: {exc}", exc_info=True)
        return {
            "schema_version": "1.0",
            "tool": "listSharePointFiles",
            "listed_at": datetime.now(timezone.utc).isoformat(),
            "error": str(exc),
            "files": [],
        }
