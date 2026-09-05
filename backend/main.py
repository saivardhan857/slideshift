"""FastAPI server — SlideShift."""

import os
import uuid
import shutil
import tempfile
import asyncio
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from parser import parse_source
from transfer import transfer
from validator import validate


TEMP_DIR = Path(tempfile.gettempdir()) / "ctm_jobs"
TEMP_DIR.mkdir(exist_ok=True)

FRONTEND_DIR = Path(__file__).parent.parent / "frontend"


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Clean up any stale temp jobs on startup
    for d in TEMP_DIR.iterdir():
        if d.is_dir():
            shutil.rmtree(d, ignore_errors=True)
    yield
    # Cleanup on shutdown
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
    if file.size and file.size > 100 * 1024 * 1024:  # 100 MB limit
        raise HTTPException(400, f"{label} exceeds 100 MB size limit.")


def _cleanup_job(job_dir: Path):
    shutil.rmtree(job_dir, ignore_errors=True)


@app.post("/api/transfer")
async def transfer_endpoint(
    background_tasks: BackgroundTasks,
    source: UploadFile = File(...),
    template: UploadFile = File(...),
):
    _validate_pptx_upload(source, "Source presentation")
    _validate_pptx_upload(template, "College template")

    job_id = str(uuid.uuid4())
    job_dir = TEMP_DIR / job_id
    job_dir.mkdir()

    source_path = job_dir / "source.pptx"
    template_path = job_dir / "template.pptx"
    output_path = job_dir / "Converted_Presentation.pptx"

    try:
        # Save uploads
        source_path.write_bytes(await source.read())
        template_path.write_bytes(await template.read())

        # Parse source
        try:
            source_slides = parse_source(str(source_path))
        except Exception as e:
            raise HTTPException(422, f"Could not read your presentation: {e}")

        if not source_slides:
            raise HTTPException(422, "Your presentation appears to be empty.")

        # Transfer
        try:
            results = transfer(
                source_slides=source_slides,
                source_path=str(source_path),
                template_path=str(template_path),
                output_path=str(output_path),
            )
        except Exception as e:
            raise HTTPException(500, f"Transfer failed: {e}")

        # Validate
        report = validate(str(output_path), len(source_slides), results)

        if not report.ok:
            raise HTTPException(500, f"Output validation failed: {'; '.join(report.messages)}")

        # Schedule cleanup after 10 minutes
        async def delayed_cleanup():
            await asyncio.sleep(600)
            _cleanup_job(job_dir)

        background_tasks.add_task(delayed_cleanup)

        content_warnings = [w for r in results for w in r.content_warnings]

        return JSONResponse({
            "job_id": job_id,
            "slides_processed": report.slide_count_source,
            "slides_output": report.slide_count_output,
            "slides_requiring_review": report.slides_with_warnings,
            "warnings": report.messages,
            "content_warnings": content_warnings,
            "summary": report.summary,
            "download_url": f"/api/download/{job_id}",
        })

    except HTTPException:
        _cleanup_job(job_dir)
        raise
    except Exception as e:
        _cleanup_job(job_dir)
        raise HTTPException(500, f"Unexpected error: {e}")


@app.get("/api/download/{job_id}")
async def download(job_id: str, background_tasks: BackgroundTasks):
    # Basic path traversal guard
    if ".." in job_id or "/" in job_id or "\\" in job_id:
        raise HTTPException(400, "Invalid job ID.")

    output_path = TEMP_DIR / job_id / "Converted_Presentation.pptx"
    if not output_path.exists():
        raise HTTPException(404, "File not found or already cleaned up. Please re-run the transfer.")

    # Clean up after download
    background_tasks.add_task(_cleanup_job, TEMP_DIR / job_id)

    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="Converted_Presentation.pptx",
    )


@app.get("/health")
async def health():
    return {"status": "ok"}


# Serve frontend
if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
