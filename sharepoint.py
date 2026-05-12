import asyncio
import json
import logging
import os
import time
from typing import Optional
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

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
    print(token_url)

    resp_data = await asyncio.to_thread(_sync_post_form, token_url, payload, 30)
    data = json.loads(resp_data.decode("utf-8"))

    _TOKEN_CACHE["token"] = data["access_token"]
    _TOKEN_CACHE["expires_at"] = now + data.get("expires_in", 3600)
    logger.info("SharePoint access token acquired / refreshed")
    return _TOKEN_CACHE["token"]


async def download_file(url: str) -> bytes:
    """
    Download a file from SharePoint using a Bearer token.
    Accepts direct SharePoint URLs or Microsoft Graph download URLs.
    """
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    try:
        return await asyncio.to_thread(_sync_get, url, headers, 120)
    except HTTPError as exc:
        if exc.code == 401:
            _TOKEN_CACHE["token"] = None
            token = await get_access_token()
            headers["Authorization"] = f"Bearer {token}"
            return await asyncio.to_thread(_sync_get, url, headers, 120)
        raise


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
