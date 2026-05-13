from fastmcp import FastMCP
import httpx
import os
import re
from urllib.parse import urlparse, unquote
from starlette.responses import JSONResponse
import uvicorn
import json
from dotenv import load_dotenv

load_dotenv()

SHAREPOINT_TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID", "")
SHAREPOINT_CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID", "")
SHAREPOINT_CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET", "")

mcp = FastMCP("verdantas-document-tools")


async def _get_graph_token() -> str:
    """Acquire a Microsoft Graph access token using client credentials."""
    token_url = f"https://login.microsoftonline.com/{SHAREPOINT_TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "grant_type": "client_credentials",
        "SHAREPOINT_CLIENT_ID": SHAREPOINT_CLIENT_ID,
        "SHAREPOINT_CLIENT_SECRET": SHAREPOINT_CLIENT_SECRET,
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
# --- Team site: /sites/SiteName/DocLib/folder/file.pdf ---
    site_match = re.match(r"/sites/([^/]+)/(.*)", path)
    if site_match:
        site_name = site_match.group(1)
        file_path = site_match.group(2)

        # Resolve site ID
        site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_name}"
        site_data_bytes = await asyncio.to_thread(_sync_get, site_url, headers, 30)
        site_data = json.loads(site_data_bytes.decode("utf-8"))
        site_id = site_data["id"]

        # Split into library name + relative file path
        # e.g. "Shared Documents/folder/file.pdf"
        #       → library="Shared Documents", rel_path="folder/file.pdf"
        path_parts = file_path.split("/", 1)
        library_name = path_parts[0]
        rel_path = path_parts[1] if len(path_parts) > 1 else ""

        # Find the drive matching the library display name
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        drives_bytes = await asyncio.to_thread(_sync_get, drives_url, headers, 30)
        drives = json.loads(drives_bytes.decode("utf-8")).get("value", [])

        drive_id = None
        for drive in drives:
            if drive.get("name", "").lower() == library_name.lower():
                drive_id = drive["id"]
                break

        if not drive_id:
            raise ValueError(
                f"Could not find drive named '{library_name}' in site '{site_name}'. "
                f"Available drives: {[d.get('name') for d in drives]}"
            )

        # Download file from the correct drive
        encoded_rel_path = quote(rel_path, safe="/")
        item_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_rel_path}:/content"
        try:
            return await asyncio.to_thread(_sync_get, item_url, headers, 120)
        except Exception as e:
            raise ValueError(
                f"Graph download failed for '{rel_path}' in drive '{library_name}': {e}"
            )

    # --- Direct path with file extension (fallback) ---
    if re.search(r"\.\w{2,5}$", path):
        encoded_path = quote(path, safe="/")
        direct_url = f"https://{hostname}{encoded_path}"
        try:
            return await asyncio.to_thread(_sync_get, direct_url, headers, 120)
        except Exception as e:
            raise ValueError(f"Direct path download failed for '{path}': {e}")

    raise ValueError(f"Could not determine download method for URL: {url[:200]}")


async def _download_via_graph(document_url: str) -> bytes:
    """Download a SharePoint/OneDrive file using Microsoft Graph API."""
    import base64

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

            # Resolve the personal site via Sites API
            site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/personal/{user_principal}"
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
            graph_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content"
            resp = await client.get(graph_url, headers=headers)
            if resp.status_code != 200:
                raise ValueError(f"Graph shares API error ({resp.status_code}): {resp.text[:500]}")
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
                raise ValueError(f"Could not resolve site '{site_path}': {site_resp.text[:300]}")
            site_id = site_resp.json()["id"]

            item_url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}"
                f"/drive/root:/{file_path}:/content"
            )
            resp = await client.get(item_url, headers=headers)
            if resp.status_code != 200:
                raise ValueError(f"Graph download error ({resp.status_code}): {resp.text[:500]}")
            return resp.content

        raise ValueError(
            f"Could not determine download method for URL type '{info['type']}'. "
            f"URL: {document_url[:200]}"
        )


@mcp.tool()
async def listSupportedFormats() -> list:
    """
    Lists the document formats supported for analysis.

    Returns:
        A list of supported file format objects with extension and description.
    """
    return [
        {"extension": ".pdf", "description": "PDF documents (text-based)"},
        {"extension": ".docx", "description": "Microsoft Word documents"},
        {"extension": ".txt", "description": "Plain text files"},
    ]


@mcp.tool()
async def analyzeDocument(document_url: str, question: str = "") -> dict:
    """
    Analyzes a document from a URL. Downloads the document, extracts text,
    and provides an AI-powered summary or answers a specific question.
    Supports SharePoint URLs (authenticated via Microsoft Graph) and public URLs.

    Args:
        document_url: The URL of the document (SharePoint or public URL)
        question: Optional question to ask about the document. If empty, returns a full summary.

    Returns:
        Analysis results including extracted text preview and AI response.
    """
    content = None

    if "sharepoint.com" in document_url:
        if not all([SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET]):
            return {
                "error": "SharePoint integration not configured. "
                         "Set SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET env vars."
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
        from io import BytesIO
        reader = PdfReader(BytesIO(content))
        text = "\n".join(
            page.extract_text() or "" for page in reader.pages
        )
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


# Health check endpoint
@mcp.custom_route("/health", methods=["GET"])
async def health_check(request):
    return JSONResponse({"status": "healthy", "service": "verdantas-mcp-server"})


# IMPORTANT: This must come AFTER all @mcp.tool() and @mcp.custom_route() definitions
app = mcp.http_app()

if __name__ == "__main__":
    port_ = int(os.environ.get("WEBSITES_PORT", 8080))
    uvicorn.run(app, host="0.0.0.0", port=port_)
