"""FastAPI server — SlideShift."""

import asyncio
import json
import shutil
import tempfile
import uuid
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


TEMP_DIR = Path(tempfile.gettempdir()) / "ctm_jobs"
TEMP_DIR.mkdir(exist_ok=True)

SAVED_TPL_PATH = Path(__file__).parent / "saved_template.pptx"
SAVED_TPL_META = Path(__file__).parent / "saved_template.json"

FRONTEND_DIR = Path(__file__).parent.parent / "frontend"


@asynccontextmanager
async def lifespan(app: FastAPI):
    for d in TEMP_DIR.iterdir():
        if d.is_dir():
            shutil.rmtree(d, ignore_errors=True)
    yield
    for d in TEMP_DIR.iterdir():
        if d.is_dir():
            shutil.rmtree(d, ignore_errors=True)


app = FastAPI(title="SlideShift", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def _validate_pptx_upload(file: UploadFile, label: str):
    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(400, f"{label} must be a .pptx file.")
    if file.size and file.size > 100 * 1024 * 1024:
        raise HTTPException(400, f"{label} exceeds 100 MB size limit.")


def _cleanup_job(job_dir: Path):
    shutil.rmtree(job_dir, ignore_errors=True)


async def _delayed_cleanup(job_dir: Path):
    await asyncio.sleep(600)
    _cleanup_job(job_dir)


# --- Saved template endpoints ---

@app.get("/api/template/info")
async def template_info():
    if SAVED_TPL_PATH.exists():
        meta = json.loads(SAVED_TPL_META.read_text()) if SAVED_TPL_META.exists() else {}
        return {"saved": True, "filename": meta.get("filename", "template.pptx")}
    return {"saved": False}


@app.post("/api/template")
async def save_template(template: UploadFile = File(...)):
    _validate_pptx_upload(template, "Template")
    SAVED_TPL_PATH.write_bytes(await template.read())
    SAVED_TPL_META.write_text(json.dumps({"filename": template.filename}))
    return {"saved": True, "filename": template.filename}


@app.delete("/api/template")
async def clear_template():
    SAVED_TPL_PATH.unlink(missing_ok=True)
    SAVED_TPL_META.unlink(missing_ok=True)
    return {"saved": False}


# --- Transfer endpoint (SSE streaming) ---

@app.post("/api/transfer")
async def transfer_endpoint(
    source: UploadFile = File(...),
    template: Optional[UploadFile] = File(None),
):
    _validate_pptx_upload(source, "Source presentation")

    # Resolve template bytes: uploaded file takes priority, then saved template
    if template is not None and template.filename:
        _validate_pptx_upload(template, "College template")
        template_bytes = await template.read()
    elif SAVED_TPL_PATH.exists():
        template_bytes = SAVED_TPL_PATH.read_bytes()
    else:
        raise HTTPException(400, "No template provided and no saved template found.")

    source_bytes = await source.read()

    job_id = str(uuid.uuid4())
    job_dir = TEMP_DIR / job_id
    job_dir.mkdir()

    source_path = job_dir / "source.pptx"
    template_path = job_dir / "template.pptx"
    output_path = job_dir / "Converted_Presentation.pptx"

    source_path.write_bytes(source_bytes)
    template_path.write_bytes(template_bytes)

    # Parse source before streaming starts so parse errors return normal HTTP errors
    try:
        source_slides = parse_source(str(source_path))
    except Exception as e:
        _cleanup_job(job_dir)
        raise HTTPException(422, f"Could not read your presentation: {e}")

    if not source_slides:
        _cleanup_job(job_dir)
        raise HTTPException(422, "Your presentation appears to be empty.")

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
            await queue.put({"type": "progress", "step": "validating"})
            report = validate(str(output_path), len(source_slides), results)
            if not report.ok:
                _cleanup_job(job_dir)
                await queue.put({
                    "type": "error",
                    "message": f"Validation failed: {'; '.join(report.messages)}",
                })
                return
            content_warnings = [w for r in results for w in r.content_warnings]
            await queue.put({
                "type": "done",
                "job_id": job_id,
                "slides_processed": report.slide_count_source,
                "slides_output": report.slide_count_output,
                "slides_requiring_review": report.slides_with_warnings,
                "warnings": report.messages,
                "content_warnings": content_warnings,
                "summary": report.summary,
                "download_url": f"/api/download/{job_id}",
            })
            asyncio.create_task(_delayed_cleanup(job_dir))
        except Exception as e:
            _cleanup_job(job_dir)
            await queue.put({"type": "error", "message": str(e)})

    asyncio.create_task(run_transfer())

    async def generate():
        # Source was already parsed — signal step 0 done, include slide count
        yield f"data: {json.dumps({'type': 'progress', 'step': 'parsed', 'total_slides': len(source_slides)})}\n\n"
        while True:
            try:
                item = await asyncio.wait_for(queue.get(), timeout=300)
            except asyncio.TimeoutError:
                _cleanup_job(job_dir)
                yield f"data: {json.dumps({'type': 'error', 'message': 'Transfer timed out.'})}\n\n"
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
    # re-click Download. _delayed_cleanup (10 min) + startup sweep collect it.
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
