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


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """Handle tool calls by dispatching to the appropriate handler."""
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
    """Run the EHS MCP Server using stdio transport."""
    async with stdio_server() as (read_stream, write_stream):
        logger.info("EHS MCP Server running on stdio")
        await app.run(read_stream, write_stream, app.create_initialization_options())


def run_sse(host: str, port: int):
    """Run the EHS MCP Server using SSE transport."""
    sse_transport = SseServerTransport("/messages")
    async def handle_sse(request):
        async with sse_transport.connect_sse(
            request.scope, request.receive, request._send
        ) as streams:
            await app.run(streams[0], streams[1], app.create_initialization_options())

    starlette_app = Starlette(
        routes=[
            Route("/sse", endpoint=handle_sse),
            Route("/messages", endpoint=sse_transport.handle_post_message),
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
        help="Transport layer (stdio for local clients, sse for HTTP/remote)",
    )
    parser.add_argument("--host", default="0.0.0.0", help="SSE host (default: 0.0.0.0)")
    parser.add_argument("--port", type=int, default=8080, help="SSE port (default: 8080)")
    args = parser.parse_args()

    if args.transport == "sse":
        run_sse(args.host, args.port)
    else:
        asyncio.run(run_stdio())
