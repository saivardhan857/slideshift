"""Thin async adapter that drives the EXISTING SlideShift HTTP routes
in-process (no second network hop, no second port) via httpx's ASGI
transport. It contains zero conversion/validation logic of its own: it
only knows how to call POST /api/transfer and GET /api/download/{id}
exactly as a real HTTP client would, and to speak the SSE protocol that
/api/transfer already emits.
"""
import json
from typing import Any, AsyncIterator, Optional

import httpx
from fastapi import FastAPI

_PPTX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


class ConversionRequestError(Exception):
    def __init__(self, message: str, status_code: int):
        super().__init__(message)
        self.status_code = status_code


def _client(app: FastAPI) -> httpx.AsyncClient:
    transport = httpx.ASGITransport(app=app)
    return httpx.AsyncClient(transport=transport, base_url="http://mcp-internal")


def _extract_detail(body: bytes) -> str:
    try:
        detail = json.loads(body).get("detail")
        if isinstance(detail, str) and detail:
            return detail
    except Exception:
        pass
    return "The request to the SlideShift conversion service failed."


async def stream_transfer(
    app: FastAPI,
    source_bytes: bytes,
    source_filename: str,
    template_bytes: Optional[bytes] = None,
    template_filename: Optional[str] = None,
) -> AsyncIterator[dict[str, Any]]:
    """Calls the existing POST /api/transfer route and yields each SSE
    event exactly as the real frontend receives it."""
    files = {"source": (source_filename, source_bytes, _PPTX_CONTENT_TYPE)}
    if template_bytes is not None:
        files["template"] = (template_filename or "template.pptx", template_bytes, _PPTX_CONTENT_TYPE)

    async with _client(app) as client:
        async with client.stream("POST", "/api/transfer", files=files) as resp:
            if resp.status_code != 200:
                body = await resp.aread()
                raise ConversionRequestError(_extract_detail(body), resp.status_code)
            async for line in resp.aiter_lines():
                if not line.startswith("data: "):
                    continue
                yield json.loads(line[len("data: "):])


async def download(app: FastAPI, slideshift_job_id: str) -> bytes:
    """Calls the existing GET /api/download/{job_id} route unchanged."""
    async with _client(app) as client:
        resp = await client.get(f"/api/download/{slideshift_job_id}")
        if resp.status_code != 200:
            raise ConversionRequestError(_extract_detail(resp.content), resp.status_code)
        return resp.content
