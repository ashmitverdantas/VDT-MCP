import asyncio
import json
import logging
import os
import time
from typing import Optional
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
import re, base64
from urllib.parse import urlparse, unquote, quote

logger = logging.getLogger(__name__)

_TOKEN_CACHE: dict = {"token": None, "expires_at": 0}

def _get_env(key: str) -> str:
    val = os.getenv(key)
    if not val:
        raise EnvironmentError(
            f"Required environment variable '{key}' is not set. "
            "Add it to your .env file or Azure App Service configuration."
        )
    return val


def _sync_post_form(url: str, form_data: dict[str, str], timeout: int = 30) -> bytes:
    request = Request(
        url,
        data=urlencode(form_data).encode("utf-8"),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with urlopen(request, timeout=timeout) as resp:
        return resp.read()


def _sync_get(url: str, headers: dict[str, str], timeout: int = 30) -> bytes:
    request = Request(url, headers=headers, method="GET")
    with urlopen(request, timeout=timeout) as resp:
        return resp.read()


async def get_access_token() -> str:
    """
    Returns a valid Bearer token for SharePoint, refreshing if expired.
    Uses the Microsoft identity platform client-credentials flow.
    """
    now = time.time()
    if _TOKEN_CACHE["token"] and now < _TOKEN_CACHE["expires_at"] - 60:
        return _TOKEN_CACHE["token"]

    tenant_id = _get_env("SHAREPOINT_TENANT_ID")
    client_id = _get_env("SHAREPOINT_CLIENT_ID")
    client_secret = _get_env("SHAREPOINT_CLIENT_SECRET")

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }

    resp_data = await asyncio.to_thread(_sync_post_form, token_url, payload, 30)
    data = json.loads(resp_data.decode("utf-8"))

    _TOKEN_CACHE["token"] = data["access_token"]
    _TOKEN_CACHE["expires_at"] = now + data.get("expires_in", 3600)
    logger.info("SharePoint access token acquired / refreshed")
    return _TOKEN_CACHE["token"]


async def download_file(url: str) -> bytes:
    """
    Download a SharePoint file via Microsoft Graph API.
    Handles team sites, personal OneDrive, and sharing links.
    """
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    parsed = urlparse(url)
    hostname = parsed.hostname or ""
    path = unquote(parsed.path)

    # --- Sharing link (/:b:/, /:w:/, /:x:/) ---
    if re.match(r"/:[a-z]:/", path):
        encoded = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
        share_id = "u!" + encoded
        graph_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content"
        try:
            return await asyncio.to_thread(_sync_get, graph_url, headers, 120)
        except Exception as e:
            raise ValueError(f"Sharing link download failed: {e}")

    # --- Team site: /sites/SiteName/DocLib/path/file.ext ---
    site_match = re.match(r"/sites/([^/]+)/(.*)", path)
    if site_match:
        site_name = site_match.group(1)
        full_file_path = site_match.group(2)

        # Resolve site ID
        site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_name}"
        site_data_bytes = await asyncio.to_thread(_sync_get, site_url, headers, 30)
        site_data = json.loads(site_data_bytes.decode("utf-8"))
        site_id = site_data["id"]

        # Split into: library name (e.g. "Shared Documents") + relative path
        # e.g. "Shared Documents/[Ext]ESS_foo.pdf" → library="Shared Documents", rel="[Ext]ESS_foo.pdf"
        path_parts = full_file_path.split("/", 1)
        library_name = path_parts[0]
        rel_path = path_parts[1] if len(path_parts) > 1 else ""

        # Find the drive whose name matches the document library
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        drives_bytes = await asyncio.to_thread(_sync_get, drives_url, headers, 30)
        drives = json.loads(drives_bytes.decode("utf-8")).get("value", [])

        # SharePoint URLs always say "Shared Documents" but the Graph API
        # drive is named "Documents" — map common aliases automatically.
        _LIBRARY_ALIASES: dict[str, list[str]] = {
            "shared documents": ["documents"],
            "documents":        ["shared documents"],
        }
        candidates = [library_name.lower()] + _LIBRARY_ALIASES.get(library_name.lower(), [])

        drive_id = None
        for drive in drives:
            if drive.get("name", "").lower() in candidates:
                drive_id = drive["id"]
                break

        if not drive_id:
            available = [d.get("name") for d in drives]
            raise ValueError(
                f"Could not find drive named '{library_name}' in site '{site_name}'. "
                f"Available drives: {available}"
            )

        # Download using the resolved drive ID
        encoded_rel_path = quote(rel_path, safe="/")
        item_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_rel_path}:/content"
        try:
            return await asyncio.to_thread(_sync_get, item_url, headers, 120)
        except Exception as e:
            raise ValueError(f"Graph download failed for '{rel_path}': {e}")

    # --- Personal OneDrive: /personal/user/Documents/... ---
    personal_match = re.match(r"/personal/([^/]+)/(.*)", path)
    if personal_match:
        user_principal = personal_match.group(1)
        file_path = personal_match.group(2)
        if file_path.startswith("Documents/"):
            file_path = file_path[len("Documents/"):]

        site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/personal/{user_principal}"
        site_data_bytes = await asyncio.to_thread(_sync_get, site_url, headers, 30)
        site_data = json.loads(site_data_bytes.decode("utf-8"))
        site_id = site_data["id"]

        encoded_file_path = quote(file_path, safe="/")
        item_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{encoded_file_path}:/content"
        try:
            return await asyncio.to_thread(_sync_get, item_url, headers, 120)
        except Exception as e:
            raise ValueError(f"Graph download failed for '{file_path}': {e}")

    raise ValueError(f"Could not determine download method for URL: {url[:200]}")
async def list_files_in_folder(
    site_id: str,
    drive_id: str,
    folder_path: str = "root",
    file_types: Optional[list[str]] = None,
) -> list[dict]:
    """
    List files in a SharePoint document library folder via Microsoft Graph.

    Args:
        site_id:     SharePoint site ID (from Graph)
        drive_id:    Document library drive ID
        folder_path: Folder path relative to drive root (default: 'root')
        file_types:  Optional list of extensions to filter, e.g. ['.pdf', '.docx']

    Returns:
        List of file metadata dicts with keys: name, id, size, url, modified, extension
    """
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    if folder_path == "root":
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root/children"
    else:
        cleaned_path = folder_path.strip("/")
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{cleaned_path}:/children"

    resp_data = await asyncio.to_thread(_sync_get, graph_url, headers, 30)
    items = json.loads(resp_data.decode("utf-8")).get("value", [])

    files = []
    lower_file_types = [ft.lower() for ft in file_types] if file_types else None
    for item in items:
        if "folder" in item:
            continue  # Skip sub-folders (not recursive by default)
        name: str = item.get("name", "")
        ext = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if lower_file_types and ext not in lower_file_types:
            continue
        files.append(
            {
                "name": name,
                "id": item.get("id"),
                "size_bytes": item.get("size", 0),
                "download_url": item.get("@microsoft.graph.downloadUrl") or item.get("webUrl"),
                "web_url": item.get("webUrl"),
                "last_modified": item.get("lastModifiedDateTime"),
                "extension": ext,
                "created_by": item.get("createdBy", {}).get("user", {}).get("displayName"),
            }
        )
    return files