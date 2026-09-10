"""FastAPI server — SlideShift."""

import asyncio
import io
import json
import logging
import shutil
import tempfile
import time
import uuid
import zipfile
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from parser import parse_source
from transfer import transfer
from validator import validate
from diagnostics import build_diagnostics

logger = logging.getLogger("slideshift")
logger.setLevel(logging.INFO)
if not logger.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s slideshift %(message)s"))
    logger.addHandler(_h)
logger.propagate = False


TEMP_DIR = Path(tempfile.gettempdir()) / "ctm_jobs"
TEMP_DIR.mkdir(exist_ok=True)

SAVED_TPL_PATH = Path(__file__).parent / "saved_template.pptx"
SAVED_TPL_META = Path(__file__).parent / "saved_template.json"
# Committed fallback so a fresh deploy always has a working template even though
# Render's free tier has no persistent disk and wipes any user-saved one.
DEFAULT_TPL_PATH = Path(__file__).parent / "default_template.pptx"
DEFAULT_TPL_NAME = "CUCOM Template.pptx"

FRONTEND_DIR = Path(__file__).parent.parent / "frontend"

# --- Phase 4 hardening limits (explicit, conservative; real lecture decks are
# a few MB and up to ~90 slides, so these leave generous headroom) ---
MAX_UPLOAD_BYTES = 100 * 1024 * 1024        # 100 MB per uploaded file
MAX_SLIDES = 500                            # reject pathological slide counts
_PPTX_REQUIRED_MEMBER = "ppt/presentation.xml"
_SSE_IDLE_TIMEOUT = 300                     # seconds with no progress event -> give up
_JOB_TTL_AFTER_DONE = 600                   # keep a finished job downloadable this long
_JOB_TTL_SAFETY_NET = 1800                  # hard upper bound for any job directory


@asynccontextmanager
async def lifespan(app: FastAPI):
    _sweep_temp()
    yield
    _sweep_temp()


def _sweep_temp():
    try:
        for d in TEMP_DIR.iterdir():
            if d.is_dir():
                shutil.rmtree(d, ignore_errors=True)
    except FileNotFoundError:
        pass


app = FastAPI(title="SlideShift", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# --------------------------------------------------------------------------- #
# Safety helpers
# --------------------------------------------------------------------------- #

def _safe_error(exc: Exception, context: str) -> str:
    """Log the real exception server-side; return a generic client-safe string.

    Never lets a filesystem path, traceback, temp-dir name or dependency
    exception dump reach the API response.
    """
    logger.exception("%s: %s", context, type(exc).__name__)
    return "An unexpected error occurred while processing your presentation."


def _check_pptx_filename(file: UploadFile, label: str):
    fn = (file.filename or "").lower()
    if not fn.endswith(".pptx"):
        logger.warning("rejected upload (%s): bad extension %r", label, file.filename)
        raise HTTPException(400, f"{label} must be a .pptx file.")


async def _read_capped(file: UploadFile, label: str) -> bytes:
    """Read an upload while enforcing MAX_UPLOAD_BYTES on the bytes actually
    received — never trusts the optional client-supplied `file.size`."""
    data = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        logger.warning("rejected upload (%s): over %d bytes", label, MAX_UPLOAD_BYTES)
        raise HTTPException(413, f"{label} exceeds the {MAX_UPLOAD_BYTES // (1024 * 1024)} MB size limit.")
    return data


def _validate_pptx_bytes(data: bytes, label: str):
    """Reject empty files and files that only *look* like .pptx by name.

    A .pptx is an OOXML ZIP container; a valid one contains ppt/presentation.xml.
    This is a lightweight structural check, not full schema validation.
    """
    if not data:
        raise HTTPException(400, f"{label} is empty.")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = set(zf.namelist())
    except zipfile.BadZipFile:
        logger.warning("rejected upload (%s): not a zip/OOXML container", label)
        raise HTTPException(400, f"{label} is not a valid PowerPoint (.pptx) file.")
    if _PPTX_REQUIRED_MEMBER not in names:
        logger.warning("rejected upload (%s): missing %s", label, _PPTX_REQUIRED_MEMBER)
        raise HTTPException(400, f"{label} is not a valid PowerPoint (.pptx) file.")


def _cleanup_job(job_dir: Path):
    shutil.rmtree(job_dir, ignore_errors=True)
    logger.debug("cleaned job dir %s", job_dir.name)


async def _delayed_cleanup(job_dir: Path, delay: int = _JOB_TTL_AFTER_DONE):
    await asyncio.sleep(delay)
    _cleanup_job(job_dir)


# --------------------------------------------------------------------------- #
# Saved template endpoints
# --------------------------------------------------------------------------- #

@app.get("/api/template/info")
async def template_info():
    if SAVED_TPL_PATH.exists():
        meta = json.loads(SAVED_TPL_META.read_text()) if SAVED_TPL_META.exists() else {}
        return {"saved": True, "filename": meta.get("filename", "template.pptx")}
    if DEFAULT_TPL_PATH.exists():
        return {"saved": True, "filename": DEFAULT_TPL_NAME, "default": True}
    return {"saved": False}


@app.post("/api/template")
async def save_template(template: UploadFile = File(...)):
    _check_pptx_filename(template, "Template")
    data = await _read_capped(template, "Template")
    _validate_pptx_bytes(data, "Template")
    SAVED_TPL_PATH.write_bytes(data)
    SAVED_TPL_META.write_text(json.dumps({"filename": template.filename}))
    logger.info("saved template %r (%d bytes)", template.filename, len(data))
    return {"saved": True, "filename": template.filename}


@app.delete("/api/template")
async def clear_template():
    SAVED_TPL_PATH.unlink(missing_ok=True)
    SAVED_TPL_META.unlink(missing_ok=True)
    return {"saved": False}


# --------------------------------------------------------------------------- #
# Transfer endpoint (SSE streaming)
# --------------------------------------------------------------------------- #

@app.post("/api/transfer")
async def transfer_endpoint(
    source: UploadFile = File(...),
    template: Optional[UploadFile] = File(None),
):
    _check_pptx_filename(source, "Source presentation")
    source_bytes = await _read_capped(source, "Source presentation")
    _validate_pptx_bytes(source_bytes, "Source presentation")

    # Resolve template bytes: uploaded file takes priority, then saved template
    if template is not None and template.filename:
        _check_pptx_filename(template, "College template")
        template_bytes = await _read_capped(template, "College template")
        _validate_pptx_bytes(template_bytes, "College template")
    elif SAVED_TPL_PATH.exists():
        template_bytes = SAVED_TPL_PATH.read_bytes()
    elif DEFAULT_TPL_PATH.exists():
        template_bytes = DEFAULT_TPL_PATH.read_bytes()
    else:
        raise HTTPException(400, "No template provided and no saved template found.")

    job_id = str(uuid.uuid4())
    job_dir = TEMP_DIR / job_id
    job_dir.mkdir()
    # Hard upper bound on this directory's lifetime, whatever happens next
    # (client disconnect, SSE timeout, crash before the normal cleanup is scheduled).
    asyncio.create_task(_delayed_cleanup(job_dir, _JOB_TTL_SAFETY_NET))

    source_path = job_dir / "source.pptx"
    template_path = job_dir / "template.pptx"
    output_path = job_dir / "Converted_Presentation.pptx"

    source_path.write_bytes(source_bytes)
    template_path.write_bytes(template_bytes)

    # Parse source before streaming starts so parse errors return normal HTTP errors
    try:
        source_slides = parse_source(str(source_path))
    except Exception as e:
        logger.warning("job %s: parse failed (%s)", job_id, type(e).__name__)
        _cleanup_job(job_dir)
        raise HTTPException(422, "Could not read your presentation. Please make sure it is a valid, uncorrupted .pptx file.")

    if not source_slides:
        _cleanup_job(job_dir)
        raise HTTPException(422, "Your presentation appears to be empty.")

    if len(source_slides) > MAX_SLIDES:
        _cleanup_job(job_dir)
        raise HTTPException(
            413,
            f"Your presentation has {len(source_slides)} slides; the maximum supported is {MAX_SLIDES}.",
        )

    logger.info("job %s: start — %d slides", job_id, len(source_slides))

    loop = asyncio.get_running_loop()
    queue: asyncio.Queue = asyncio.Queue()

    def sync_progress(step: str):
        if step.startswith("slide:"):
            event = {"type": "progress", "step": "transferring", "slide": step[6:]}
        else:
            event = {"type": "progress", "step": step}
        asyncio.run_coroutine_threadsafe(queue.put(event), loop)

    async def run_transfer():
        try:
            t0 = time.monotonic()
            results = await loop.run_in_executor(
                None,
                lambda: transfer(
                    source_slides=source_slides,
                    source_path=str(source_path),
                    template_path=str(template_path),
                    output_path=str(output_path),
                    progress_callback=sync_progress,
                ),
            )
            t_transfer = time.monotonic() - t0

            await queue.put({"type": "progress", "step": "validating"})
            t1 = time.monotonic()
            report = validate(str(output_path), len(source_slides), results)
            t_validate = time.monotonic() - t1

            if not report.ok:
                logger.warning("job %s: validation failed — %s", job_id, "; ".join(report.messages))
                _cleanup_job(job_dir)
                await queue.put({
                    "type": "error",
                    "message": "The converted presentation failed validation and was not produced. Please try again.",
                })
                return

            diag = build_diagnostics(
                results, report,
                transfer_seconds=t_transfer, validation_seconds=t_validate,
            )
            content_warnings = [w for r in results for w in r.content_warnings]
            logger.info(
                "job %s: done — %d/%d slides, %d review, %.1fs",
                job_id, diag.slides_output, diag.slides_processed,
                diag.slides_requiring_review, diag.total_seconds,
            )
            await queue.put({
                "type": "done",
                "job_id": job_id,
                "slides_processed": report.slide_count_source,
                "slides_output": report.slide_count_output,
                "slides_requiring_review": report.slides_with_warnings,
                "warnings": report.messages,
                "content_warnings": content_warnings,
                "summary": report.summary,
                "diagnostics": diag.to_dict(),
                "download_url": f"/api/download/{job_id}",
            })
            asyncio.create_task(_delayed_cleanup(job_dir, _JOB_TTL_AFTER_DONE))
        except Exception as e:
            _cleanup_job(job_dir)
            await queue.put({"type": "error", "message": _safe_error(e, f"job {job_id}: transfer")})

    asyncio.create_task(run_transfer())

    async def generate():
        # Source was already parsed — signal step 0 done, include slide count
        yield f"data: {json.dumps({'type': 'progress', 'step': 'parsed', 'total_slides': len(source_slides)})}\n\n"
        while True:
            try:
                item = await asyncio.wait_for(queue.get(), timeout=_SSE_IDLE_TIMEOUT)
            except asyncio.TimeoutError:
                # Do NOT delete the job dir here — the worker thread may still be
                # running. The safety-net _delayed_cleanup + startup sweep collect it.
                logger.warning("job %s: SSE idle timeout", job_id)
                yield f"data: {json.dumps({'type': 'error', 'message': 'The transfer is taking longer than expected. Please try again.'})}\n\n"
                break
            yield f"data: {json.dumps(item)}\n\n"
            if item["type"] in ("done", "error"):
                break

    return StreamingResponse(generate(), media_type="text/event-stream")


@app.get("/api/download/{job_id}")
async def download(job_id: str):
    if ".." in job_id or "/" in job_id or "\\" in job_id:
        raise HTTPException(400, "Invalid job ID.")
    output_path = TEMP_DIR / job_id / "Converted_Presentation.pptx"
    if not output_path.exists():
        raise HTTPException(404, "File not found or already cleaned up. Please re-run the transfer.")
    # No cleanup here — the browser may issue the request twice and users
    # re-click Download. _delayed_cleanup + startup sweep collect it.
    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="Converted_Presentation.pptx",
    )


@app.get("/health")
async def health():
    return {"status": "ok"}


if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
