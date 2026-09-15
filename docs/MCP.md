# SlideShift MCP

An [MCP](https://modelcontextprotocol.io) (Model Context Protocol) interface
that lets an MCP-compatible AI client drive SlideShift's existing PPTX
conversion pipeline. It is purely additive: the web app, its API, and the
conversion engine are unchanged.

## Architecture

```
MCP Client
    |
    | HTTP (Streamable HTTP transport) + Authorization: Bearer <key>
    v
POST/GET/DELETE /mcp    (backend/mcp_integration/, mounted on the same
    |                     FastAPI app, same process, same Render service)
    v
POST /api/transfer, GET /api/download/{id}   (unchanged, existing routes)
    |
    v
parser.py / transfer.py / validator.py / ai_classify.py   (unchanged)
```

The MCP layer never parses, transfers, or validates a presentation itself.
Every tool either:

- calls the existing `/api/transfer` or `/api/download/{id}` routes
  in-process, via an `httpx.ASGITransport` bound to the running FastAPI app
  (`mcp_integration/adapter.py`) — this runs the real route handlers, not a
  reimplementation, or
- reads data those calls already produced, held in a small in-memory job
  registry (`mcp_integration/state.py`).

`mcp_integration/` is a self-contained package: `auth.py` (bearer-token
middleware), `state.py` (job registry), `adapter.py` (calls into the
existing routes), `tools.py` (tool/resource definitions), `server.py`
(wires it all onto `/mcp`). `backend/main.py` only gained one import, one
`app.mount`-equivalent wiring call, and a two-line change to its existing
`lifespan` so the MCP session manager starts and stops with the app.

## Available tools

| Tool | Purpose |
| --- | --- |
| `convert_presentation` | Submit a source (and optional template) `.pptx` as base64. Returns `{success, job_id, status: "queued"}` immediately; the real conversion runs via the existing `/api/transfer` route in the background. |
| `get_conversion_status` | Poll a job started with `convert_presentation`. `status` mirrors the existing SSE step names: `queued, parsed, classifying, transferring, validating, done, error`. |
| `validate_presentation` | Returns the validation report SlideShift already produced during the transfer (slide counts, warnings, errors, diagnostics) — no re-validation, no duplicated rules. |
| `get_converted_file` | Fetches the finished `.pptx` as base64-encoded bytes via the existing `/api/download/{id}` route. Never exposes a filesystem path. |

Resource: `slideshift://about` — a short, read-only description of
SlideShift and this workflow.

Typical flow: `convert_presentation` → poll `get_conversion_status` until
`done`/`error` → `validate_presentation` → `get_converted_file`.

## Authentication

Every request to `/mcp` requires:

```
Authorization: Bearer <SLIDESHIFT_MCP_KEY>
```

checked with a constant-time comparison. Behavior:

- `SLIDESHIFT_MCP_KEY` set → that exact token is required; missing, wrong,
  or malformed headers get `401`.
- `SLIDESHIFT_MCP_KEY` unset → the endpoint refuses all requests (`503`)
  **unless** `SLIDESHIFT_MCP_ALLOW_UNAUTHENTICATED=true` is also set. That
  flag is for local development only — never set it in production.

Auth is isolated ASGI middleware wrapping only `/mcp`
(`mcp_integration/auth.py`); it cannot affect the existing frontend, CORS,
or API routes.

## Local development

```
cd backend
pip install -r requirements.txt

# generate a local key (PowerShell):
# [Convert]::ToBase64String((1..32 | ForEach-Object { Get-Random -Maximum 256 }))

$env:SLIDESHIFT_MCP_KEY = "<paste the generated key>"
uvicorn main:app --host 0.0.0.0 --port 8765
```

The existing web app still runs unchanged at `http://localhost:8765/`. MCP
is available at `http://localhost:8765/mcp`.

To try it without generating a key first (dev only):

```
$env:SLIDESHIFT_MCP_ALLOW_UNAUTHENTICATED = "true"
uvicorn main:app --host 0.0.0.0 --port 8765
```

## Production configuration (Render)

Add one environment variable to the existing `slideshift` Render service
(Dashboard → your service → Environment):

```
SLIDESHIFT_MCP_KEY = <a long random string, generated as shown above>
```

Do **not** set `SLIDESHIFT_MCP_ALLOW_UNAUTHENTICATED` in production. No
other Render configuration changes — the start command, `rootDir`, and
build command in `render.yaml` are unchanged; MCP runs in the same
process as the existing app.

Endpoint once deployed:

```
https://slideshift.onrender.com/mcp
```

## Example MCP client configuration

Exact config keys vary by client. A typical `mcp.json`-style entry for an
HTTP/Streamable-HTTP MCP client:

```json
{
  "mcpServers": {
    "slideshift": {
      "url": "https://slideshift.onrender.com/mcp",
      "headers": {
        "Authorization": "Bearer <SLIDESHIFT_MCP_KEY>"
      }
    }
  }
}
```

Never commit the real key. Substitute it from your client's own secret
storage / environment variable mechanism.

## Security considerations

- **No filesystem exposure.** Tools never accept or return a filesystem
  path; `get_converted_file` returns base64 bytes obtained through the
  existing download route.
- **No SSRF.** No tool accepts or fetches a URL.
- **Job IDs are opaque and validated by lookup, not parsed.** An unknown
  or path-traversal-shaped `job_id` simply misses the in-memory registry
  and returns a generic `JOB_NOT_FOUND` — it's never used to build a path.
- **Errors are sanitized.** Tool errors reuse the same user-safe messages
  the existing `/api/transfer` route already returns to the web frontend;
  raw exceptions, stack traces, and paths are never surfaced (`main.py`'s
  existing `_safe_error` behavior is unchanged and still applies to the
  underlying routes MCP calls).
- **Existing upload limits apply.** MCP submissions go through the same
  `/api/transfer` route, so the existing 100 MB upload cap and `.pptx`
  structural validation apply unchanged.
- **DNS-rebinding host/origin checks are intentionally disabled** in the
  MCP SDK's transport (`mcp_integration/server.py`) because bearer-token
  auth already gates every request before that check would run; a spoofed
  `Host` header cannot bypass it.

## Testing

```
cd backend
python test_mcp_integration.py    # auth matrix, tool discovery, conversion
                                   # end-to-end, security, no regressions
python -m pytest -q               # existing suite, unaffected
python test_transfer.py           # existing self-contained scripts, unaffected
python test_qa_regression.py
# ...and the rest of backend/test_*.py, per backend/README.md
```

## Troubleshooting

- **`503 MCP_AUTH_UNCONFIGURED`** — `SLIDESHIFT_MCP_KEY` isn't set and
  `SLIDESHIFT_MCP_ALLOW_UNAUTHENTICATED` isn't `true`. Set one of them.
- **`401 MCP_UNAUTHORIZED`** — the `Authorization: Bearer <key>` header is
  missing, malformed, or doesn't match `SLIDESHIFT_MCP_KEY`.
- **`JOB_NOT_FOUND`** — the job id is unknown or has expired. Job records
  are kept in memory for 600s after completion (300s after an error),
  mirroring the existing job directory TTL in `main.py`; convert again.
- **`convert_presentation` never reaches `done`** — check
  `get_conversion_status`'s `error` field once `status` becomes `error`;
  it's the same message the web app would show for that failure (bad
  `.pptx`, no template available, etc.).
- **A fresh Render deploy has no saved template** — this is expected and
  unrelated to MCP: `/api/transfer` (and therefore `convert_presentation`
  without `template_base64`) falls back to the bundled
  `backend/default_template.pptx`, same as the web app.
