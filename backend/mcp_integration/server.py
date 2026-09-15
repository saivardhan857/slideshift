"""Builds the SlideShift MCP server and its mountable ASGI app.

Call build_mcp_app(app) once, with the already-constructed FastAPI `app`
(its routes don't need to be registered yet -- the adapter calls into it
lazily, at request time). Returns (asgi_app, session_manager): mount
asgi_app at "/mcp" and fold session_manager.run() into main.py's existing
lifespan, or the streamable HTTP transport never actually processes
requests (see docs/MCP.md).
"""
from fastapi import FastAPI
from mcp.server.fastmcp import FastMCP
from mcp.server.streamable_http_manager import StreamableHTTPSessionManager
from mcp.server.transport_security import TransportSecuritySettings

from .auth import BearerAuthMiddleware
from .tools import register_tools

_INSTRUCTIONS = (
    "SlideShift converts a source PowerPoint into a college template's "
    "format. Start with convert_presentation, poll get_conversion_status, "
    "then validate_presentation and get_converted_file."
)


class _PathRewriteASGI:
    """Forwards to `app` with the scope's HTTP path replaced by `path`.

    A Mount strips its prefix before forwarding, so "/mcp/" reaches the
    inner app's own "/" route -- but an exact-match Route does not rewrite
    anything, so the bare "/mcp" (no trailing slash, which is how MCP
    clients conventionally address the endpoint) needs this to reach that
    same route too.
    """

    def __init__(self, app, path: str):
        self.app = app
        self.path = path

    async def __call__(self, scope, receive, send) -> None:
        if scope["type"] == "http":
            scope = {**scope, "path": self.path, "raw_path": self.path.encode("utf-8")}
        await self.app(scope, receive, send)


def build_mcp_app(app: FastAPI) -> StreamableHTTPSessionManager:
    """Wires the MCP tools onto /mcp of the given FastAPI app and returns
    the session manager whose .run() must be entered in the app's
    lifespan."""
    mcp = FastMCP(
        name="slideshift",
        instructions=_INSTRUCTIONS,
        streamable_http_path="/",  # mounted at "/mcp" below -> final path is "/mcp"
        stateless_http=True,       # each tool call is independent; job state lives in state.py
        json_response=True,        # plain JSON responses -- no client-side SSE parsing needed
        # The SDK's own Host/Origin (DNS-rebinding) check would otherwise run
        # for every request; it is redundant here because BearerAuthMiddleware
        # below already gates ALL access on the shared secret first -- a
        # spoofed Host header cannot bypass that check.
        transport_security=TransportSecuritySettings(enable_dns_rebinding_protection=False),
    )
    register_tools(mcp, app)

    mcp_asgi_app = BearerAuthMiddleware(mcp.streamable_http_app())
    # Mount only matches "/mcp/..." (Starlette requires the trailing slash);
    # add_route covers the bare "/mcp" that MCP clients actually use.
    app.add_route(
        "/mcp", _PathRewriteASGI(mcp_asgi_app, "/"),
        methods=["GET", "POST", "DELETE"], include_in_schema=False,
    )
    app.mount("/mcp", mcp_asgi_app)
    return mcp.session_manager
