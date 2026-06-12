"""
FastAPI application for the ICAI NCE Conversion Tool.

Routes:
  GET  /                 -> single-page UI
  POST /api/extract      -> upload files + constitution -> structured JSON (Claude)
  POST /api/generate     -> structured JSON -> formula-linked .xlsx download
  GET  /health           -> health check

The /api/generate endpoint works WITHOUT an API key (pure engine), so the manual
or review-then-generate path needs no Claude. /api/extract requires ANTHROPIC_API_KEY.
"""
from __future__ import annotations
import io, json, os
from dataclasses import asdict
from typing import List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Header
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from models import Payload
from builder import build_workbook

app = FastAPI(title="ICAI NCE Conversion Tool", version="1.0.0")


def valid_codes():
    """Access codes are managed via the ACCESS_CODES env var (comma-separated),
    shared with the genius-tb-tool. Default 'genius2025' for parity."""
    raw = os.environ.get("ACCESS_CODES", "genius2025")
    return {c.strip() for c in raw.split(",") if c.strip()}


def require_code(code):
    if not code or code.strip() not in valid_codes():
        raise HTTPException(401, "Invalid or inactive access code. Please contact "
                                 "M S Joshi & Co. (connect@msjc.in) for access.")
    return code.strip()

HERE = os.path.dirname(__file__)
STATIC_DIR = HERE
if os.path.isdir(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/health")
def health():
    return {"status": "ok", "llm_configured": bool(os.environ.get("ANTHROPIC_API_KEY")), "codes_configured": bool(os.environ.get("ACCESS_CODES"))}


@app.get("/", response_class=HTMLResponse)
def index():
    path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(path):
        return HTMLResponse(open(path, encoding="utf-8").read())
    return HTMLResponse("<h1>ICAI NCE Conversion Tool</h1><p>UI not found.</p>")


@app.post("/api/verify")
def api_verify(x_access_code: str = Header(None)):
    require_code(x_access_code)
    return {"valid": True}


@app.post("/api/extract")
async def api_extract(constitution: str = Form(...),
                      extra: str = Form(""),
                      files: List[UploadFile] = File(...),
                      x_access_code: str = Header(None)):
    require_code(x_access_code)
    if constitution not in ("proprietorship", "partnership"):
        raise HTTPException(400, "constitution must be 'proprietorship' or 'partnership'")
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise HTTPException(503, "ANTHROPIC_API_KEY is not configured on the server. "
                                 "Use the manual JSON path, or set the key and redeploy.")
    payload_files = []
    for f in files:
        nm = (f.filename or "").lower()
        if nm.endswith(".doc") and not nm.endswith(".docx"):
            raise HTTPException(415, "This looks like a legacy .doc file (old Word format). "
                                     "Please open it in Word and use File > Save As > Word Document (.docx), "
                                     "then upload the .docx. PDF and Excel are also supported.")
        payload_files.append((f.filename, await f.read()))
    try:
        from llm import extract_payload, reconcile_and_correct
        import reconcile as rec
        payload = extract_payload(constitution, payload_files, extra)
        payload, discrepancies, fixes = reconcile_and_correct(constitution, payload_files, payload)
    except Exception as e:
        raise HTTPException(500, f"Extraction failed: {e}")
    return JSONResponse({
        "payload": asdict(payload),
        "reconciliation": {
            "passed": len(discrepancies) == 0,
            "discrepancies": discrepancies,
            "fixes": fixes,
            "report": rec.report_text(discrepancies),
        },
    })


@app.post("/api/generate")
async def api_generate(payload: dict, x_access_code: str = Header(None)):
    require_code(x_access_code)
    import reconcile as rec
    try:
        p = Payload.parse(payload)
    except Exception as e:
        raise HTTPException(400, f"Invalid payload: {e}")
    discrepancies = rec.reconcile(p)
    if discrepancies:
        # ZERO TOLERANCE: never ship a non-reconciling workbook
        raise HTTPException(status_code=422, detail={
            "error": "RECONCILIATION_FAILED",
            "message": "The figures do not reconcile to the source. No workbook produced.",
            "discrepancies": discrepancies,
            "report": rec.report_text(discrepancies),
        })
    try:
        data = build_workbook(p)
    except Exception as e:
        raise HTTPException(400, f"Build failed: {e}")
    name = (p.entity.name or "Entity").replace(" ", "_")
    fy = p.entity.cy_fy or "FY"
    fname = f"{name}_FS_FY{fy}_ICAI_NCE.xlsx"
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )
