"""EHS MCP Server entry point.

Provides CLI arguments to run the server using either stdio or SSE transport.
"""
from dotenv import load_dotenv
import argparse
import asyncio
import json
import logging

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

# ── MCP Server ────────────────────────────────────────────────────────────────

app = Server("ehs-mcp-server")

TOOLS: list[Tool] = [
    ANALYZE_DOCUMENTS_TOOL,
    LIST_FILES_TOOL,
    GET_METADATA_TOOL,
]

HANDLERS = {
    "analyzeDocuments": handle_analyze_documents,
    "listSharePointFiles": handle_list_sharepoint_files,
    "getDocumentMetadata": handle_get_document_metadata,
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
            text=json.dumps(result, indent=2, ensure_ascii=False)
        )
    ]


# ── Transport ─────────────────────────────────────────────────────────────────

async def run_stdio():
    async with stdio_server() as (read_stream, write_stream):
        logger.info("EHS MCP Server running on stdio")
        await app.run(read_stream, write_stream, app.create_initialization_options())


def run_sse(host: str, port: int):
    sse_transport = SseServerTransport("/messages/")

    # Pure ASGI class — bypasses Starlette routing entirely.
    # Starlette always wraps endpoint functions into Request objects,
    # which strips the 'send' callable. A plain ASGI class avoids this.
    class ASGIApp:
        async def __call__(self, scope, receive, send):
            if scope["type"] != "http":
                return
            path = scope.get("path", "")
            if path == "/sse":
                async with sse_transport.connect_sse(scope, receive, send) as streams:
                    await app.run(
                        streams[0], streams[1],
                        app.create_initialization_options()
                    )
            elif path.startswith("/messages"):
                await sse_transport.handle_post_message(scope, receive, send)
            else:
                await send({
                    "type": "http.response.start",
                    "status": 404,
                    "headers": [[b"content-type", b"text/plain"]],
                })
                await send({
                    "type": "http.response.body",
                    "body": b"Not found",
                })

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
