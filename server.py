"""EHS MCP Server entry point.

Provides CLI arguments to run the server using either stdio or SSE transport.
"""

import argparse
import asyncio
import json
import logging

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.server.sse import SseServerTransport
from mcp.types import Tool, TextContent
from starlette.applications import Starlette
from starlette.routing import Route
import uvicorn

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

# ── Server bootstrap ──────────────────────────────────────────────────────────

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


# ── Transport selection ───────────────────────────────────────────────────────

async def run_stdio():
    async with stdio_server() as (read_stream, write_stream):
        logger.info("EHS MCP Server running on stdio")
        await app.run(read_stream, write_stream, app.create_initialization_options())


def run_sse(host: str, port: int):
    sse_transport = SseServerTransport("/messages")

    # ✅ Both routes must be raw ASGI functions (scope, receive, send)
    # NOT Starlette Request objects — that's what caused the 'no attribute send' error

    async def handle_sse(scope, receive, send):
        """SSE endpoint — raw ASGI, not a Starlette Request handler."""
        async with sse_transport.connect_sse(scope, receive, send) as streams:
            await app.run(
                streams[0], streams[1], app.create_initialization_options()
            )

    async def handle_messages(scope, receive, send):
        """Message POST endpoint — raw ASGI."""
        await sse_transport.handle_post_message(scope, receive, send)

    starlette_app = Starlette(
        routes=[
            Route("/sse", endpoint=handle_sse),
            Route("/messages", endpoint=handle_messages, methods=["POST"]),
        ]
    )

    logger.info("EHS MCP Server running on SSE at http://%s:%s/sse", host, port)
    uvicorn.run(starlette_app, host=host, port=port)


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="EHS MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
    )
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8080)
    args = parser.parse_args()

    if args.transport == "sse":
        run_sse(args.host, args.port)
    else:
        asyncio.run(run_stdio())
