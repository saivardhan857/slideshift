"""MCP tool implementations. Every tool either calls the existing HTTP
routes via adapter.py, or reads results those calls already produced
(state.py) -- no PPTX parsing/transfer/validation logic lives here.
"""
import asyncio
import base64
import binascii
import logging
import uuid
from typing import Any, Optional

from fastapi import FastAPI

from . import state
from .adapter import ConversionRequestError, download, stream_transfer

logger = logging.getLogger("slideshift")

# Best-effort client-side guard mirroring main.py's MAX_UPLOAD_BYTES (100 MB).
# The existing /api/transfer route remains the single source of truth for
# this limit -- this just avoids doing base64 decode work on hopeless input.
_MAX_UPLOAD_BYTES = 100 * 1024 * 1024
_DONE_EXPIRY_SECONDS = 600  # matches main.py's _JOB_TTL_AFTER_DONE
_ERROR_EXPIRY_SECONDS = 300


def _decode_b64(data: str, label: str) -> bytes:
    if not data:
        raise ValueError(f"{label} is required.")
    try:
        decoded = base64.b64decode(data, validate=True)
    except (binascii.Error, ValueError) as e:
        raise ValueError(f"{label} is not valid base64.") from e
    if not decoded:
        raise ValueError(f"{label} is empty.")
    if len(decoded) > _MAX_UPLOAD_BYTES:
        raise ValueError(f"{label} exceeds the {_MAX_UPLOAD_BYTES // (1024 * 1024)} MB size limit.")
    return decoded


def _check_pptx_name(filename: str, label: str) -> None:
    if not filename.lower().endswith(".pptx"):
        raise ValueError(f"{label} must name a .pptx file.")


async def _run_conversion(
    app: FastAPI,
    job_id: str,
    source_bytes: bytes,
    source_filename: str,
    template_bytes: Optional[bytes],
    template_filename: Optional[str],
) -> None:
    job = state.get(job_id)
    if job is None:
        return
    try:
        async for event in stream_transfer(app, source_bytes, source_filename, template_bytes, template_filename):
            etype = event.get("type")
            if etype == "progress":
                job.status = event.get("step", job.status)
                job.slide = event.get("slide")
            elif etype == "done":
                job.status = "done"
                job.slideshift_job_id = event.get("job_id")
                job.download_url = event.get("download_url")
                job.result = event
                state.schedule_expiry(job_id, _DONE_EXPIRY_SECONDS)
            elif etype == "error":
                job.status = "error"
                job.error = event.get("message", "Conversion failed.")
                state.schedule_expiry(job_id, _ERROR_EXPIRY_SECONDS)
    except ConversionRequestError as e:
        job.status = "error"
        job.error = str(e)
        state.schedule_expiry(job_id, _ERROR_EXPIRY_SECONDS)
    except Exception as e:
        logger.exception("mcp job %s: unexpected failure (%s)", job_id, type(e).__name__)
        job.status = "error"
        job.error = "An unexpected error occurred while processing your presentation."
        state.schedule_expiry(job_id, _ERROR_EXPIRY_SECONDS)


def register_tools(mcp, app: FastAPI) -> None:
    """Registers all SlideShift MCP tools/resources onto `mcp` (a FastMCP
    instance), bound to the running FastAPI `app` they call into."""

    @mcp.tool()
    async def convert_presentation(
        source_base64: str,
        source_filename: str = "presentation.pptx",
        template_base64: Optional[str] = None,
        template_filename: Optional[str] = None,
    ) -> dict[str, Any]:
        """Convert a source PowerPoint into the configured college template
        using SlideShift's existing conversion pipeline.

        source_base64: base64-encoded .pptx bytes of the presentation to
            convert.
        template_base64: optional base64-encoded .pptx bytes of the
            template to use. If omitted, SlideShift falls back to the
            saved/default template, exactly as the web app does.

        Returns immediately with a job_id; poll get_conversion_status(job_id)
        for progress, then validate_presentation / get_converted_file.
        """
        try:
            _check_pptx_name(source_filename, "source_filename")
            source_bytes = _decode_b64(source_base64, "source_base64")
            template_bytes = None
            if template_base64:
                _check_pptx_name(template_filename or "template.pptx", "template_filename")
                template_bytes = _decode_b64(template_base64, "template_base64")
        except ValueError as e:
            return {"success": False, "error": str(e), "code": "INVALID_INPUT"}

        job_id = str(uuid.uuid4())
        state.create(job_id)
        asyncio.create_task(_run_conversion(
            app, job_id, source_bytes, source_filename, template_bytes, template_filename,
        ))
        return {"success": True, "job_id": job_id, "status": "queued"}

    @mcp.tool()
    def get_conversion_status(job_id: str) -> dict[str, Any]:
        """Get the current status of a conversion started with
        convert_presentation. status is one of: queued, parsed,
        classifying, transferring, validating, done, error."""
        job = state.get(job_id)
        if job is None:
            return {"success": False, "error": "Unknown or expired job ID.", "code": "JOB_NOT_FOUND"}
        payload: dict[str, Any] = {"success": True, "job_id": job_id, "status": job.status}
        if job.slide:
            payload["slide"] = job.slide
        if job.status == "error":
            payload["error"] = job.error
        if job.status == "done" and job.result:
            payload.update({
                "slides_processed": job.result.get("slides_processed"),
                "slides_output": job.result.get("slides_output"),
                "slides_requiring_review": job.result.get("slides_requiring_review"),
                "summary": job.result.get("summary"),
            })
        return payload

    @mcp.tool()
    def validate_presentation(job_id: str) -> dict[str, Any]:
        """Return SlideShift's own validation report for a finished
        conversion job (slide counts, warnings, errors, diagnostics) --
        produced by the same validation pass the web app already ran
        during conversion. Does not re-run or duplicate validation."""
        job = state.get(job_id)
        if job is None:
            return {"success": False, "error": "Unknown or expired job ID.", "code": "JOB_NOT_FOUND"}
        if job.status == "error":
            return {
                "success": False, "job_id": job_id, "status": "error",
                "error": job.error, "code": "CONVERSION_ERROR",
            }
        if job.status != "done" or not job.result:
            return {
                "success": False, "job_id": job_id, "status": job.status,
                "error": "Conversion has not finished yet.", "code": "JOB_NOT_READY",
            }
        r = job.result
        return {
            "success": True,
            "job_id": job_id,
            "slide_count_source": r.get("slides_processed"),
            "slide_count_output": r.get("slides_output"),
            "slides_requiring_review": r.get("slides_requiring_review"),
            "warnings": r.get("warnings", []),
            "content_warnings": r.get("content_warnings", []),
            "summary": r.get("summary"),
            "diagnostics": r.get("diagnostics"),
        }

    @mcp.tool()
    async def get_converted_file(job_id: str) -> dict[str, Any]:
        """Fetch the converted .pptx for a finished job as base64-encoded
        bytes, via SlideShift's existing download route. Never exposes a
        filesystem path."""
        job = state.get(job_id)
        if job is None:
            return {"success": False, "error": "Unknown or expired job ID.", "code": "JOB_NOT_FOUND"}
        if job.status != "done" or not job.slideshift_job_id:
            return {
                "success": False, "job_id": job_id, "status": job.status,
                "error": "Conversion has not finished yet.", "code": "JOB_NOT_READY",
            }
        try:
            content = await download(app, job.slideshift_job_id)
        except ConversionRequestError as e:
            return {"success": False, "job_id": job_id, "error": str(e), "code": "DOWNLOAD_FAILED"}
        return {
            "success": True,
            "job_id": job_id,
            "filename": "Converted_Presentation.pptx",
            "size_bytes": len(content),
            "content_base64": base64.b64encode(content).decode("ascii"),
        }

    @mcp.resource("slideshift://about")
    def about() -> str:
        """Read-only description of SlideShift and its MCP tools."""
        return (
            "SlideShift converts a source PowerPoint (.pptx) into a college "
            "template's format, preserving text, images and tables.\n\n"
            "Workflow: convert_presentation(source_base64, ...) -> job_id, "
            "then poll get_conversion_status(job_id) until status is "
            "'done' or 'error', then validate_presentation(job_id) for the "
            "validation report and get_converted_file(job_id) for the "
            "resulting .pptx (base64-encoded).\n\n"
            "If no template is supplied, SlideShift falls back to the "
            "saved or default template, the same as the web app."
        )
