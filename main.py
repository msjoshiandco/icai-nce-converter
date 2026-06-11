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

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from models import Payload
from builder import build_workbook

app = FastAPI(title="ICAI NCE Conversion Tool", version="1.0.0")

HERE = os.path.dirname(__file__)
STATIC_DIR = HERE
if os.path.isdir(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/health")
def health():
    return {"status": "ok", "llm_configured": bool(os.environ.get("ANTHROPIC_API_KEY"))}


@app.get("/", response_class=HTMLResponse)
def index():
    path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(path):
        return HTMLResponse(open(path, encoding="utf-8").read())
    return HTMLResponse("<h1>ICAI NCE Conversion Tool</h1><p>UI not found.</p>")


@app.post("/api/extract")
async def api_extract(constitution: str = Form(...),
                      extra: str = Form(""),
                      files: List[UploadFile] = File(...)):
    if constitution not in ("proprietorship", "partnership"):
        raise HTTPException(400, "constitution must be 'proprietorship' or 'partnership'")
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise HTTPException(503, "ANTHROPIC_API_KEY is not configured on the server. "
                                 "Use the manual JSON path, or set the key and redeploy.")
    payload_files = []
    for f in files:
        payload_files.append((f.filename, await f.read()))
    try:
        from llm import extract_payload
        payload = extract_payload(constitution, payload_files, extra)
    except Exception as e:
        raise HTTPException(500, f"Extraction failed: {e}")
    return JSONResponse(asdict(payload))


@app.post("/api/generate")
async def api_generate(payload: dict):
    try:
        p = Payload.parse(payload)
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
