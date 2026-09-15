"""Bearer-token authentication for the MCP HTTP interface (see docs/MCP.md).

Isolated ASGI middleware that only wraps the /mcp sub-app -- it cannot
affect existing routes, CORS, or auth behavior elsewhere in the app.
"""
import hmac
import os
from typing import Optional

from starlette.datastructures import Headers
from starlette.responses import JSONResponse
from starlette.types import ASGIApp, Receive, Scope, Send


def _load_config() -> tuple[Optional[str], bool]:
    key = os.getenv("SLIDESHIFT_MCP_KEY", "").strip()
    allow_unauthenticated = os.getenv(
        "SLIDESHIFT_MCP_ALLOW_UNAUTHENTICATED", ""
    ).strip().lower() in ("1", "true", "yes")
    return (key or None), allow_unauthenticated


def _extract_bearer(header_value: Optional[str]) -> Optional[str]:
    if not header_value or not header_value.startswith("Bearer "):
        return None
    token = header_value[len("Bearer "):].strip()
    return token or None


class BearerAuthMiddleware:
    """Rejects any /mcp request without a valid `Authorization: Bearer <key>`.

    - SLIDESHIFT_MCP_KEY set -> that exact token is required (constant-time
      comparison).
    - SLIDESHIFT_MCP_KEY unset -> access is refused UNLESS
      SLIDESHIFT_MCP_ALLOW_UNAUTHENTICATED=true is also set. This is meant
      for local development only; production must set SLIDESHIFT_MCP_KEY.
    """

    def __init__(self, app: ASGIApp):
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        key, allow_unauthenticated = _load_config()

        if key is None:
            if allow_unauthenticated:
                await self.app(scope, receive, send)
                return
            response = JSONResponse(
                {
                    "success": False,
                    "error": "MCP authentication is not configured on this server.",
                    "code": "MCP_AUTH_UNCONFIGURED",
                },
                status_code=503,
            )
            await response(scope, receive, send)
            return

        headers = Headers(scope=scope)
        supplied = _extract_bearer(headers.get("authorization"))
        if supplied is None or not hmac.compare_digest(supplied, key):
            response = JSONResponse(
                {
                    "success": False,
                    "error": "Missing or invalid MCP credentials.",
                    "code": "MCP_UNAUTHORIZED",
                },
                status_code=401,
                headers={"WWW-Authenticate": "Bearer"},
            )
            await response(scope, receive, send)
            return

        await self.app(scope, receive, send)
