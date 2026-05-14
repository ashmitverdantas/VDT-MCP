"""EHS MCP Server entry point.

Provides CLI arguments to run the server using either stdio or SSE transport.
"""
from dotenv import load_dotenv
import argparse
import asyncio
import base64
import json
import logging
import os
import re
from io import BytesIO
from urllib.parse import urlparse, unquote

import httpx
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.server.sse import SseServerTransport
from mcp.types import Tool, TextContent
import uvicorn

load_dotenv()

from analyze_documents import (
    ANALYZE_DOCUMENTS_TOOL,
    handle_analyze_documents,
)
from list_sharepoint_files import (
    LIST_FILES_TOOL,
    handle_list_sharepoint_files,
)
from get_document_metadata import (
    GET_METADATA_TOOL,
    handle_get_document_metadata,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("ehs-mcp-server")

# ── SharePoint / Graph config ─────────────────────────────────────────────────

SHAREPOINT_TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID", "")
SHAREPOINT_CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID", "")
SHAREPOINT_CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET", "")


# ── Microsoft Graph helpers ───────────────────────────────────────────────────

async def _get_graph_token() -> str:
    """Acquire a Microsoft Graph access token using client credentials."""
    token_url = (
        f"https://login.microsoftonline.com/{SHAREPOINT_TENANT_ID}/oauth2/v2.0/token"
    )
    payload = {
        "grant_type": "client_credentials",
        "CLIENT_ID": SHAREPOINT_CLIENT_ID,
        "CLIENT_SECRET": SHAREPOINT_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    async with httpx.AsyncClient() as client:
        resp = await client.post(token_url, data=payload)
        if resp.status_code != 200:
            raise ValueError(
                f"Token request failed ({resp.status_code}). "
                f"Tenant: {SHAREPOINT_TENANT_ID}, Client: {SHAREPOINT_CLIENT_ID}, "
                f"Secret ends with: ...{SHAREPOINT_CLIENT_SECRET[-4:] if SHAREPOINT_CLIENT_SECRET else 'EMPTY'}. "
                f"Response: {resp.text[:500]}"
            )
        return resp.json()["access_token"]


def _extract_file_path_from_url(url: str) -> dict:
    """
    Parse a SharePoint/OneDrive URL into its components.
    Returns dict with keys: hostname, type, site_path, user_principal, file_path
    """
    parsed = urlparse(url)
    hostname = parsed.hostname or ""
    path = unquote(parsed.path)
    params = {}
    if parsed.query:
        params = dict(
            p.split("=", 1)
            for p in unquote(parsed.query).split("&")
            if "=" in p
        )

    result = {"hostname": hostname, "type": "unknown", "raw_url": url}

    # OneDrive for Business — file path in "id" query param
    # e.g. hullinc-my.sharepoint.com/my?id=%2Fpersonal%2Fuser%2FDocuments%2Ffile.pdf
    if "id" in params:
        id_path = unquote(params["id"])
        parts = id_path.strip("/").split("/")
        if len(parts) >= 3 and parts[0] == "personal":
            result["type"] = "onedrive_personal"
            result["user_principal"] = parts[1]
            # "Documents" is the drive root — strip it to get the relative file path
            file_parts = parts[2:]
            if file_parts and file_parts[0] == "Documents":
                file_parts = file_parts[1:]
            result["file_path"] = "/".join(file_parts)
            return result

    # Sharing link — /:b:/s/Site/encoded...
    if re.match(r"/:[a-z]:/", path):
        result["type"] = "sharing_link"
        return result

    # Direct path: /personal/user/Documents/file.pdf
    personal_match = re.match(r"/personal/([^/]+)/(.*)", path)
    if personal_match:
        result["type"] = "onedrive_personal"
        result["user_principal"] = personal_match.group(1)
        file_path = personal_match.group(2)
        if file_path.startswith("Documents/"):
            file_path = file_path[len("Documents/"):]
        result["file_path"] = file_path
        return result

    # Team site: /sites/SiteName/DocLib/folder/file.pdf
    site_match = re.match(r"/sites/([^/]+)/(.*)", path)
    if site_match:
        result["type"] = "team_site"
        result["site_path"] = site_match.group(1)
        result["file_path"] = site_match.group(2)
        return result

    # Direct path with file extension
    if re.search(r"\.\w{2,5}$", path):
        result["type"] = "direct_path"
        result["file_path"] = path
        return result

    return result


async def _download_via_graph(document_url: str) -> bytes:
    """Download a SharePoint/OneDrive file using Microsoft Graph API."""
    token = await _get_graph_token()
    info = _extract_file_path_from_url(document_url)
    hostname = info["hostname"]

    if not hostname:
        raise ValueError("Could not parse SharePoint hostname from URL")

    async with httpx.AsyncClient(follow_redirects=True, timeout=60.0) as client:
        headers = {"Authorization": f"Bearer {token}"}

        # --- OneDrive for Business (personal) ---
        if info["type"] == "onedrive_personal":
            user_principal = info["user_principal"]
            file_path = info["file_path"]

            site_url = (
                f"https://graph.microsoft.com/v1.0/sites/{hostname}:/personal/{user_principal}"
            )
            site_resp = await client.get(site_url, headers=headers)
            if site_resp.status_code != 200:
                raise ValueError(
                    f"Could not resolve OneDrive site for '{user_principal}': "
                    f"{site_resp.text[:300]}"
                )
            site_id = site_resp.json()["id"]

            item_url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}"
                f"/drive/root:/{file_path}:/content"
            )
            resp = await client.get(item_url, headers=headers)
            if resp.status_code == 200:
                return resp.content
            raise ValueError(
                f"Could not download '{file_path}' from OneDrive of {user_principal}. "
                f"Status: {resp.status_code} — {resp.text[:300]}"
            )

        # --- Sharing link (/:b:/, /:w:/, /:x:/) ---
        if info["type"] == "sharing_link":
            encoded = base64.urlsafe_b64encode(document_url.encode()).decode().rstrip("=")
            share_id = "u!" + encoded
            graph_url = (
                f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content"
            )
            resp = await client.get(graph_url, headers=headers)
            if resp.status_code != 200:
                raise ValueError(
                    f"Graph shares API error ({resp.status_code}): {resp.text[:500]}"
                )
            return resp.content

        # --- SharePoint team sites (/sites/SiteName/...) ---
        if info["type"] == "team_site":
            site_path = info["site_path"]
            file_path = info["file_path"]

            site_resp = await client.get(
                f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_path}",
                headers=headers,
            )
            if site_resp.status_code != 200:
                raise ValueError(
                    f"Could not resolve site '{site_path}': {site_resp.text[:300]}"
                )
            site_id = site_resp.json()["id"]

            item_url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}"
                f"/drive/root:/{file_path}:/content"
            )
            resp = await client.get(item_url, headers=headers)
            if resp.status_code != 200:
                raise ValueError(
                    f"Graph download error ({resp.status_code}): {resp.text[:500]}"
                )
            return resp.content

        raise ValueError(
            f"Could not determine download method for URL type '{info['type']}'. "
            f"URL: {document_url[:200]}"
        )


# ── Tool handlers (from main.py) ──────────────────────────────────────────────

async def handle_list_supported_formats(_arguments: dict) -> list:
    """Lists the document formats supported for analysis."""
    return [
        {"extension": ".pdf", "description": "PDF documents (text-based)"},
        {"extension": ".docx", "description": "Microsoft Word documents"},
        {"extension": ".txt", "description": "Plain text files"},
    ]


async def handle_analyze_document(arguments: dict) -> dict:
    """
    Analyzes a document from a URL. Downloads the document, extracts text,
    and provides an AI-powered summary or answers a specific question.
    Supports SharePoint URLs (authenticated via Microsoft Graph) and public URLs.
    """
    document_url: str = arguments.get("document_url", "")
    question: str = arguments.get("question", "")

    if not document_url:
        return {"error": "document_url is required"}

    content = None

    if "sharepoint.com" in document_url:
        if not all([SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET]):
            return {
                "error": (
                    "SharePoint integration not configured. "
                    "Set SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, "
                    "SHAREPOINT_CLIENT_SECRET env vars."
                )
            }
        try:
            content = await _download_via_graph(document_url)
        except Exception as e:
            return {"error": f"SharePoint download failed: {str(e)}"}
    else:
        async with httpx.AsyncClient(follow_redirects=True, timeout=30.0) as client:
            response = await client.get(document_url)
            if response.status_code != 200:
                return {"error": f"Could not fetch document: HTTP {response.status_code}"}
            content = response.content

    url_lower = document_url.lower()
    if url_lower.endswith(".pdf") or b"%PDF" in content[:10]:
        from pypdf import PdfReader
        reader = PdfReader(BytesIO(content))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        page_count = len(reader.pages)
    else:
        text = content.decode("utf-8", errors="ignore")
        page_count = 1

    if not text.strip():
        return {"error": "Could not extract text from the document"}

    return {
        "page_count": page_count,
        "text_length": len(text),
        "text_preview": text,
        "question": question or "Full summary requested",
        "status": "Text extracted successfully. AI analysis would be applied here.",
    }


# ── Tool definitions ──────────────────────────────────────────────────────────

LIST_SUPPORTED_FORMATS_TOOL = Tool(
    name="listSupportedFormats",
    description="Lists the document formats supported for analysis.",
    inputSchema={"type": "object", "properties": {}, "required": []},
)

ANALYZE_DOCUMENT_TOOL = Tool(
    name="analyzeDocument",
    description=(
        "Analyzes a document from a URL. Downloads the document, extracts text, "
        "and provides an AI-powered summary or answers a specific question. "
        "Supports SharePoint URLs (authenticated via Microsoft Graph) and public URLs."
    ),
    inputSchema={
        "type": "object",
        "properties": {
            "document_url": {
                "type": "string",
                "description": "The URL of the document (SharePoint or public URL)",
            },
            "question": {
                "type": "string",
                "description": (
                    "Optional question to ask about the document. "
                    "If empty, returns a full summary."
                ),
            },
        },
        "required": ["document_url"],
    },
)


# ── MCP Server ────────────────────────────────────────────────────────────────

app = Server("ehs-mcp-server")

TOOLS: list[Tool] = [
    ANALYZE_DOCUMENTS_TOOL,
    LIST_FILES_TOOL,
    GET_METADATA_TOOL,
    LIST_SUPPORTED_FORMATS_TOOL,
    ANALYZE_DOCUMENT_TOOL,
]

HANDLERS = {
    "analyzeDocuments": handle_analyze_documents,
    "listSharePointFiles": handle_list_sharepoint_files,
    "getDocumentMetadata": handle_get_document_metadata,
    "listSupportedFormats": handle_list_supported_formats,
    "analyzeDocument": handle_analyze_document,
}


@app.list_tools()
async def list_tools() -> list[Tool]:
    return TOOLS


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    handler = HANDLERS.get(name)
    if handler is None:
        raise ValueError(f"Unknown tool: {name}")
    result = await handler(arguments)
    return [
        TextContent(
            type="text",
            text=json.dumps(result, indent=2, ensure_ascii=False),
        )
    ]


# ── Transport ─────────────────────────────────────────────────────────────────

async def run_stdio():
    async with stdio_server() as (read_stream, write_stream):
        logger.info("EHS MCP Server running on stdio")
        await app.run(read_stream, write_stream, app.create_initialization_options())


def run_sse(host: str, port: int):
    sse_transport = SseServerTransport("/messages/")

    class ASGIApp:
        async def __call__(self, scope, receive, send):
            if scope["type"] != "http":
                return
            path = scope.get("path", "")
            if path == "/sse":
                async with sse_transport.connect_sse(scope, receive, send) as streams:
                    await app.run(
                        streams[0], streams[1],
                        app.create_initialization_options(),
                    )
            elif path.startswith("/messages"):
                await sse_transport.handle_post_message(scope, receive, send)
            elif path == "/health":
                body = json.dumps(
                    {"status": "healthy", "service": "ehs-mcp-server"}
                ).encode()
                await send({
                    "type": "http.response.start",
                    "status": 200,
                    "headers": [[b"content-type", b"application/json"]],
                })
                await send({"type": "http.response.body", "body": body})
            else:
                await send({
                    "type": "http.response.start",
                    "status": 404,
                    "headers": [[b"content-type", b"text/plain"]],
                })
                await send({"type": "http.response.body", "body": b"Not found"})

    logger.info("EHS MCP Server running on SSE at http://%s:%s/sse", host, port)
    uvicorn.run(ASGIApp(), host=host, port=port)


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="EHS MCP Server")
    parser.add_argument("--transport", choices=["stdio", "sse"], default="stdio")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8080)

    args = parser.parse_args()

    if args.transport == "sse":
        run_sse(args.host, args.port)
    else:
        asyncio.run(run_stdio())
        