"""MCP interface tests -- authentication, tool discovery, conversion,
status, validation, and security. Exercises the /mcp routes in-process
against the real FastAPI app (same TestClient pattern as
test_default_template.py), so every conversion goes through the actual,
unmodified SlideShift pipeline.

Self-contained -- no pytest.  Run:  python test_mcp_integration.py
"""
import io
import json
import os
import sys
from pathlib import Path

from fastapi.testclient import TestClient
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).parent))
import main

_passed = _failed = 0
_KEY = "unit-test-key-do-not-use-in-prod"
_HEADERS = {"Content-Type": "application/json", "Accept": "application/json, text/event-stream"}


def check(name, cond, detail=""):
    global _passed, _failed
    if cond:
        _passed += 1
        print(f"  [OK]   {name}")
    else:
        _failed += 1
        print(f"  [FAIL] {name}  {detail}")


def _pptx_bytes(n_slides=2):
    prs = Presentation()
    for i in range(n_slides):
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = f"Slide {i + 1}"
        s.placeholders[1].text_frame.text = "point one\npoint two"
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _rpc(client, method, params, headers=None):
    body = {"jsonrpc": "2.0", "id": 1, "method": method, "params": params}
    h = dict(_HEADERS)
    if headers:
        h.update(headers)
    return client.post("/mcp", json=body, headers=h)


def _call_tool(client, name, args, headers=None):
    r = _rpc(client, "tools/call", {"name": name, "arguments": args}, headers)
    if r.status_code != 200:
        return None, r
    data = r.json()
    if "error" in data:
        return None, data["error"]
    return json.loads(data["result"]["content"][0]["text"]), r


def _auth(key):
    return {"Authorization": f"Bearer {key}"}


def run():
    os.environ.pop("SLIDESHIFT_MCP_ALLOW_UNAUTHENTICATED", None)
    leaked_paths = []

    with TestClient(main.app) as client:
        # ---- Authentication -------------------------------------------- #
        os.environ.pop("SLIDESHIFT_MCP_KEY", None)
        r = _rpc(client, "initialize", {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "t", "version": "1"}})
        check("no key configured -> request rejected (not silently open)", r.status_code in (401, 503), f"got {r.status_code}")

        os.environ["SLIDESHIFT_MCP_KEY"] = _KEY

        r = _rpc(client, "initialize", {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "t", "version": "1"}})
        check("missing Authorization header -> 401", r.status_code == 401, str(r.status_code))

        r = _rpc(client, "initialize", {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "t", "version": "1"}},
                 headers=_auth("wrong-key"))
        check("wrong key -> 401", r.status_code == 401, str(r.status_code))

        r = _rpc(client, "initialize", {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "t", "version": "1"}},
                 headers={"Authorization": _KEY})  # no "Bearer " prefix
        check("malformed Authorization header -> 401", r.status_code == 401, str(r.status_code))

        r = _rpc(client, "initialize", {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "t", "version": "1"}},
                 headers=_auth(_KEY))
        check("valid key -> 200", r.status_code == 200, str(r.status_code))

        good = _auth(_KEY)

        # ---- Tool / resource discovery ----------------------------------- #
        r = _rpc(client, "tools/list", {}, headers=good)
        tool_names = {t["name"] for t in r.json()["result"]["tools"]} if r.status_code == 200 else set()
        check("all 4 tools discoverable", tool_names == {
            "convert_presentation", "get_conversion_status", "validate_presentation", "get_converted_file",
        }, str(tool_names))
        for t in r.json()["result"]["tools"]:
            check(f"{t['name']} has a valid input schema", t.get("inputSchema", {}).get("type") == "object", str(t.get("inputSchema")))

        r = _rpc(client, "resources/list", {}, headers=good)
        resource_uris = {res["uri"] for res in r.json()["result"]["resources"]} if r.status_code == 200 else set()
        check("slideshift://about resource discoverable", "slideshift://about" in resource_uris, str(resource_uris))

        # ---- Security: bad input rejected without touching the pipeline -- #
        result, _ = _call_tool(client, "convert_presentation",
                                {"source_base64": "aGVsbG8=", "source_filename": "not_a_pptx.exe"}, headers=good)
        check("non-.pptx filename rejected", result is not None and result["success"] is False and result["code"] == "INVALID_INPUT", str(result))

        result, _ = _call_tool(client, "convert_presentation",
                                {"source_base64": "!!!not-base64!!!", "source_filename": "x.pptx"}, headers=good)
        check("malformed base64 rejected", result is not None and result["success"] is False and result["code"] == "INVALID_INPUT", str(result))

        for bad_id in ["nonexistent-job", "../../../etc/passwd", "..\\..\\windows\\system32", "'; DROP TABLE jobs;--"]:
            result, _ = _call_tool(client, "get_conversion_status", {"job_id": bad_id}, headers=good)
            check(f"unknown/traversal job_id safely rejected: {bad_id!r}",
                  result is not None and result["success"] is False and result["code"] == "JOB_NOT_FOUND", str(result))
            result, _ = _call_tool(client, "get_converted_file", {"job_id": bad_id}, headers=good)
            check(f"get_converted_file rejects unknown job_id: {bad_id!r}",
                  result is not None and result["success"] is False and result["code"] == "JOB_NOT_FOUND", str(result))

        # ---- Conversion: invokes the EXISTING pipeline, end to end ------- #
        source = _pptx_bytes(3)
        import base64
        result, _ = _call_tool(client, "convert_presentation", {
            "source_base64": base64.b64encode(source).decode(),
            "source_filename": "deck.pptx",
        }, headers=good)
        check("convert_presentation returns success + job_id immediately", result is not None and result.get("success") and result.get("status") == "queued", str(result))
        job_id = result["job_id"] if result else None

        status = None
        if job_id:
            import time
            for _ in range(30):
                status, _ = _call_tool(client, "get_conversion_status", {"job_id": job_id}, headers=good)
                if status and status["status"] in ("done", "error"):
                    break
                time.sleep(0.2)
        check("conversion reaches 'done' via the real pipeline", status is not None and status["status"] == "done", str(status))

        if status and status["status"] == "done":
            v, _ = _call_tool(client, "validate_presentation", {"job_id": job_id}, headers=good)
            check("validate_presentation reuses the real validation report", v is not None and v["success"] and v["slide_count_output"] == 3, str(v))

            f, _ = _call_tool(client, "get_converted_file", {"job_id": job_id}, headers=good)
            check("get_converted_file returns base64 content", f is not None and f.get("success") and f.get("content_base64"), str(f)[:200])
            if f and f.get("success"):
                content = base64.b64decode(f["content_base64"])
                prs = Presentation(io.BytesIO(content))
                check("downloaded file is a valid, openable PPTX with the right slide count", len(prs.slides) == 3, str(len(prs.slides)))
                check("no filesystem path in get_converted_file response", "\\" not in json.dumps(f) and ":" not in f.get("filename", ""), str(f.get("filename")))

        # ---- Conversion: unsupported/corrupt file rejected safely -------- #
        result, _ = _call_tool(client, "convert_presentation", {
            "source_base64": base64.b64encode(b"not a real zip/pptx").decode(),
            "source_filename": "corrupt.pptx",
        }, headers=good)
        bad_job_id = result["job_id"] if result and result.get("success") else None
        if bad_job_id:
            import time
            bad_status = None
            for _ in range(15):
                bad_status, _ = _call_tool(client, "get_conversion_status", {"job_id": bad_job_id}, headers=good)
                if bad_status and bad_status["status"] in ("done", "error"):
                    break
                time.sleep(0.2)
            check("corrupt .pptx safely errors out (not crash, not silently 'done')",
                  bad_status is not None and bad_status["status"] == "error", str(bad_status))
            if bad_status:
                err = bad_status.get("error", "")
                leaked_paths.append(err)
                check("corrupt-file error message has no traceback/path leakage",
                      "Traceback" not in err and "C:\\" not in err and "/tmp/" not in err, err)

        # ---- Global secret-leak scan over every response we saw --------- #
        check("the configured MCP key itself is never echoed back in any response",
              _KEY not in "".join(leaked_paths), "leaked in an error message")

    os.environ.pop("SLIDESHIFT_MCP_KEY", None)
    print(f"\nMCP integration: {_passed} passed, {_failed} failed")
    return _failed == 0


if __name__ == "__main__":
    sys.exit(0 if run() else 1)
