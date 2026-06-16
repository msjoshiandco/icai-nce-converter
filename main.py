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
import io, json, os, urllib.request, urllib.parse
from dataclasses import asdict
from typing import List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Header
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from models import Payload
from builder import build_workbook

app = FastAPI(title="ICAI NCE Conversion Tool", version="1.0.0")


# In ACCESS_CODES each entry is "code" or "code:limit" (limit 0/absent = unlimited),
# e.g. "MSjoshi@725:0,Demo@12345:3". Codes must not contain a ':'.
def parse_codes():
    raw = os.environ.get("ACCESS_CODES", "genius2025")
    out = {}
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        if ":" in part:
            code, _, lim = part.rpartition(":")
            code = code.strip()
            try:
                limit = int(lim.strip())
            except ValueError:
                code, limit = part, 0
        else:
            code, limit = part, 0
        if code:
            out[code] = max(0, limit)
    return out


def valid_codes():
    return set(parse_codes().keys())


def require_code(code):
    if not code or code.strip() not in valid_codes():
        raise HTTPException(401, "Invalid or inactive access code. Please contact "
                                 "M S Joshi & Co. (connect@msjc.in) for access.")
    return code.strip()


# ---- usage counter (persistent via Upstash Redis if configured; else in-memory) ----
_MEM_USAGE = {}


def _upstash():
    url = os.environ.get("UPSTASH_REDIS_REST_URL")
    tok = os.environ.get("UPSTASH_REDIS_REST_TOKEN")
    return (url.rstrip("/"), tok) if url and tok else None


def _redis(path):
    cfg = _upstash()
    if not cfg:
        return None
    url, tok = cfg
    req = urllib.request.Request(f"{url}/{path}", headers={"Authorization": f"Bearer {tok}"})
    with urllib.request.urlopen(req, timeout=6) as r:
        return json.loads(r.read().decode()).get("result")


def _key(code):
    return "nceusage:" + urllib.parse.quote(code, safe="")


def usage_get(code):
    if _upstash():
        try:
            v = _redis("get/" + _key(code))
            return int(v) if v not in (None, "") else 0
        except Exception:
            pass
    return _MEM_USAGE.get(code, 0)


def usage_incr(code):
    if _upstash():
        try:
            return int(_redis("incr/" + _key(code)))
        except Exception:
            pass
    _MEM_USAGE[code] = _MEM_USAGE.get(code, 0) + 1
    return _MEM_USAGE[code]


def code_status(code):
    limit = parse_codes().get(code, 0)
    used = usage_get(code)
    return {"limit": limit, "used": used,
            "remaining": (None if limit <= 0 else max(0, limit - used)),
            "unlimited": limit <= 0}

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
    code = require_code(x_access_code)
    return {"valid": True, **code_status(code)}


@app.post("/api/admin/reset")
async def api_admin_reset(target: str = Form(...), x_admin_key: str = Header(None)):
    admin = os.environ.get("ADMIN_KEY")
    if not admin or x_admin_key != admin:
        raise HTTPException(401, "Admin key required.")
    if _upstash():
        try:
            _redis("del/" + _key(target))
        except Exception:
            pass
    _MEM_USAGE.pop(target, None)
    return {"reset": target, "used": usage_get(target)}


@app.post("/api/extract")
async def api_extract(constitution: str = Form(...),
                      extra: str = Form(""),
                      files: List[UploadFile] = File(...),
                      x_access_code: str = Header(None)):
    code = require_code(x_access_code)
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
    # enforce per-code conversion limit (extraction is the token-cost step)
    st = code_status(code)
    if not st["unlimited"] and st["used"] >= st["limit"]:
        raise HTTPException(429, f"This access code has reached its limit of {st['limit']} "
                                 "conversion(s). Please contact M S Joshi & Co. "
                                 "(connect@msjc.in) to renew or upgrade.")
    usage_incr(code)
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
async def api_generate(payload: dict, denomination: str = "actual", x_access_code: str = Header(None)):
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
        data = build_workbook(p, denomination)
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
