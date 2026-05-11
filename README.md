# EHS MCP Server — SharePoint Document Extractor

A **Model Context Protocol (MCP) server** that connects to SharePoint, extracts text from documents of any format, and returns structured JSON. Designed for the EHS project and compatible with any MCP-capable AI client.

---

## Features

| Capability | Detail |
|---|---|
| **Formats supported** | PDF, Word (.docx/.doc), Excel (.xlsx/.xls), Images (JPG/PNG/GIF/TIFF — OCR), Plain text, CSV |
| **Bulk processing** | 1–10 files per request |
| **Data source** | Microsoft SharePoint via Graph API (OAuth2) |
| **Transports** | `stdio` (local clients) · `SSE/HTTP` (remote agents) |
| **Compatible clients** | Claude (Anthropic) · Copilot Studio · LangChain · any MCP client |

---

## Tools Exposed

### 1. `analyzeDocuments` ⭐ Core tool
Downloads and extracts text from one or more SharePoint files.

**Input:**
```json
{
  "documentUrls": [
    "https://contoso.sharepoint.com/sites/EHS/Shared Documents/incident-2026-05.pdf",
    "https://contoso.sharepoint.com/sites/EHS/Shared Documents/checklist.xlsx"
  ],
  "includeRawText": true
}
```

**Output:**
```json
{
  "schema_version": "1.0",
  "tool": "analyzeDocuments",
  "processed_at": "2026-05-11T10:30:00Z",
  "summary": {
    "total": 2,
    "succeeded": 2,
    "failed": 0,
    "elapsed_seconds": 4.21
  },
  "documents": [
    {
      "url": "https://...incident-2026-05.pdf",
      "status": "success",
      "format": "pdf",
      "file_size_bytes": 204800,
      "content_hash": "abc123...",
      "text": "Incident Report\nDate: May 9, 2026\n...",
      "text_length": 3842,
      "error": null
    },
    {
      "url": "https://...checklist.xlsx",
      "status": "success",
      "format": "xlsx",
      "file_size_bytes": 51200,
      "content_hash": "def456...",
      "text": "=== Sheet: Checklist ===\nItem\tStatus\n...",
      "text_length": 1205,
      "error": null
    }
  ]
}
```

---

### 2. `listSharePointFiles`
List files in a document library folder.

**Input:**
```json
{
  "siteId": "contoso.sharepoint.com,abc,xyz",
  "driveId": "b!xxxxxxx",
  "folderPath": "Incident Reports/2026",
  "fileTypes": [".pdf", ".docx"]
}
```

---

### 3. `getDocumentMetadata`
Quick inspection: format, size, hash — without extracting text.

**Input:**
```json
{ "documentUrl": "https://contoso.sharepoint.com/.../report.pdf" }
```

---

## Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure SharePoint credentials
```bash
cp .env.example .env
# Edit .env with your Azure AD and SharePoint details
```

**Azure AD app registration requirements:**
- Permissions: `Sites.Read.All`, `Files.Read.All` (Microsoft Graph, application permissions)
- Admin consent granted

### 3. Run (stdio — for local Claude Desktop / Copilot Studio)
```bash
python server.py
```

### 4. Run (SSE — for remote agents / HTTP access)
```bash
python server.py --transport sse --port 8080
```

---

## Integration Guide

### Claude Desktop (`claude_desktop_config.json`)
```json
{
  "mcpServers": {
    "ehs-sharepoint": {
      "command": "python",
      "args": ["/path/to/ehs-mcp-server/server.py"],
      "env": {
        "SHAREPOINT_TENANT_ID": "...",
        "SHAREPOINT_CLIENT_ID": "...",
        "SHAREPOINT_CLIENT_SECRET": "..."
      }
    }
  }
}
```

### Copilot Studio
1. Deploy server with `--transport sse` to Azure App Service
2. In Copilot Studio → **Topics → Actions → Add an action → Model Context Protocol**
3. Enter: `https://your-app.azurewebsites.net/sse`
4. The three tools appear automatically

### LangChain / Python agent
```python
from langchain_mcp_adapters.client import MultiServerMCPClient

client = MultiServerMCPClient({
    "ehs": {
        "url": "http://localhost:8080/sse",
        "transport": "sse",
    }
})

tools = await client.get_tools()
# Tools: analyzeDocuments, listSharePointFiles, getDocumentMetadata
```

### Direct HTTP (SSE mode)
```bash
curl -X POST http://localhost:8080/messages \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "method": "tools/call",
    "params": {
      "name": "analyzeDocuments",
      "arguments": {
        "documentUrls": ["https://contoso.sharepoint.com/..."]
      }
    }
  }'
```

---

## Azure Deployment

```bash
# 1. Create App Service (Python 3.11, Linux)
az webapp create --resource-group EHS-RG --plan EHS-Plan \
  --name ehs-mcp-server --runtime "PYTHON:3.11"

# 2. Set environment variables
az webapp config appsettings set --name ehs-mcp-server \
  --resource-group EHS-RG --settings \
  SHAREPOINT_TENANT_ID="..." \
  SHAREPOINT_CLIENT_ID="..." \
  SHAREPOINT_CLIENT_SECRET="..."

# 3. Startup command (install Tesseract for OCR)
az webapp config set --name ehs-mcp-server --resource-group EHS-RG \
  --startup-file "apt-get install -y tesseract-ocr && python server.py --transport sse --port 8000"
```

---

## Project Structure

```
ehs-mcp-server/
├── server.py                    # Entry point — transport selection
├── requirements.txt
├── .env.example
├── auth/
│   └── sharepoint.py            # OAuth2 token acquisition + file download
├── tools/
│   ├── analyze_documents.py     # analyzeDocuments tool
│   ├── list_sharepoint_files.py # listSharePointFiles tool
│   └── get_document_metadata.py # getDocumentMetadata tool
├── processors/
│   └── extractors.py            # Per-format text extractors
└── utils/
    └── format_detection.py      # URL + magic-byte format detection
```

---

## Performance Targets (from EHS tickets)

| Scenario | Target |
|---|---|
| Single file | < 30 seconds |
| 5 files | < 2 minutes |
| 10 files | < 5 minutes |

Files are downloaded concurrently (`asyncio.gather`) to meet these targets.

---

## Security Notes

- Credentials are read from environment variables only — never hard-coded
- SharePoint access uses Microsoft's OAuth2 client-credentials flow (no user impersonation)
- Files are processed in memory and never written to disk
- Error messages never expose internal credentials or stack traces to clients
- All inter-service communication uses HTTPS/TLS
