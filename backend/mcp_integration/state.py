"""In-memory registry for MCP-initiated conversion jobs.

This mirrors, rather than replaces, the existing SSE job lifecycle in
main.py. It exists only because an MCP tool call needs to return
immediately with a job id and let later calls poll for the result, whereas
the existing frontend just stays subscribed to one SSE stream for the
whole conversion. Every record maps 1:1 to a real SlideShift job created
by the existing /api/transfer route.
"""
import asyncio
import time
from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class McpJob:
    job_id: str
    status: str = "queued"  # queued|parsed|classifying|transferring|validating|done|error
    slide: Optional[str] = None
    slideshift_job_id: Optional[str] = None
    download_url: Optional[str] = None
    result: Optional[dict[str, Any]] = None
    error: Optional[str] = None
    created_at: float = field(default_factory=time.monotonic)


_jobs: dict[str, McpJob] = {}


def create(job_id: str) -> McpJob:
    job = McpJob(job_id=job_id)
    _jobs[job_id] = job
    return job


def get(job_id: str) -> Optional[McpJob]:
    return _jobs.get(job_id)


def schedule_expiry(job_id: str, delay: float) -> None:
    async def _later():
        await asyncio.sleep(delay)
        _jobs.pop(job_id, None)

    asyncio.create_task(_later())
