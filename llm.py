"""
LLM extraction layer. Sends the source documents (PDF/Excel/images, converted to
text or passed as images) to Claude with the constitution-specific system prompt,
and parses the returned JSON into a models.Payload.

Requires env var ANTHROPIC_API_KEY at runtime. Network access required.
"""
from __future__ import annotations
import os, json, base64, re
from typing import List, Tuple

from models import Payload
from prompts import system_prompt

DEFAULT_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")


def _client():
    import anthropic  # imported lazily so the engine works without the package
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        raise RuntimeError("ANTHROPIC_API_KEY is not set")
    return anthropic.Anthropic(api_key=key)


def _content_blocks(files: List[Tuple[str, bytes]]) -> list:
    """Turn uploaded files into Claude content blocks.
    PDFs → document blocks; images → image blocks; text/csv/xlsx-text → text."""
    blocks = []
    for name, data in files:
        ext = name.lower().rsplit(".", 1)[-1] if "." in name else ""
        if ext == "pdf":
            blocks.append({
                "type": "document",
                "source": {"type": "base64", "media_type": "application/pdf",
                           "data": base64.b64encode(data).decode()},
                "title": name,
            })
        elif ext in ("png", "jpg", "jpeg", "gif", "webp"):
            mt = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
            blocks.append({
                "type": "image",
                "source": {"type": "base64", "media_type": mt,
                           "data": base64.b64encode(data).decode()},
            })
        elif ext in ("xlsx", "xls"):
            blocks.append({"type": "text", "text": f"[Spreadsheet {name}]\n" + _xlsx_to_text(data)})
        else:  # txt, csv, md, docx-text fallback
            try:
                blocks.append({"type": "text", "text": f"[File {name}]\n" + data.decode("utf-8", "ignore")})
            except Exception:
                pass
    return blocks


def _xlsx_to_text(data: bytes) -> str:
    import io, openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"## Sheet: {ws.title}")
        for row in ws.iter_rows(values_only=True):
            cells = [("" if v is None else str(v)) for v in row]
            if any(cells):
                out.append("\t".join(cells))
    return "\n".join(out)


def extract_payload(constitution: str, files: List[Tuple[str, bytes]],
                    extra_instructions: str = "") -> Payload:
    """Call Claude and return a parsed Payload."""
    client = _client()
    blocks = _content_blocks(files)
    user_text = ("Extract and classify the following source statements into the "
                 "JSON contract. Return JSON only.\n")
    if extra_instructions:
        user_text += "\nAdditional instructions from the user:\n" + extra_instructions + "\n"
    blocks.insert(0, {"type": "text", "text": user_text})

    msg = client.messages.create(
        model=DEFAULT_MODEL,
        max_tokens=8000,
        system=system_prompt(constitution),
        messages=[{"role": "user", "content": blocks}],
    )
    text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    data = _extract_json(text)
    data.setdefault("entity", {})["constitution"] = constitution
    return Payload.parse(data)


def _extract_json(text: str) -> dict:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
    # find the outermost JSON object
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("No JSON object found in model output")
    return json.loads(text[start:end + 1])


# --------------------------------------------------------------------------
# Reconciliation-driven correction loop
# --------------------------------------------------------------------------
def correct_payload(constitution, files, current_json: str, discrepancies_text: str):
    """Re-prompt Claude with the failing payload + discrepancies; return corrected Payload."""
    from prompts import CORRECTION_INSTRUCTION
    client = _client()
    blocks = _content_blocks(files)
    user_text = CORRECTION_INSTRUCTION.format(discrepancies=discrepancies_text,
                                              current_json=current_json)
    blocks.insert(0, {"type": "text", "text": user_text})
    msg = client.messages.create(
        model=DEFAULT_MODEL, max_tokens=8000,
        system=system_prompt(constitution),
        messages=[{"role": "user", "content": blocks}],
    )
    text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    data = _extract_json(text)
    data.setdefault("entity", {})["constitution"] = constitution
    return Payload.parse(data)


def reconcile_and_correct(constitution, files, payload, max_retries: int = 2):
    """Run the reconciliation engine; auto-fix; retry via Claude up to max_retries.
    Returns (payload, discrepancies, fixes_applied)."""
    import json as _json
    from dataclasses import asdict
    import reconcile as rec

    fixes = []
    discr = rec.reconcile(payload)
    if discr:
        fixes += rec.auto_fix(payload)
        discr = rec.reconcile(payload)

    attempts = 0
    while discr and attempts < max_retries:
        try:
            payload = correct_payload(constitution, files,
                                      _json.dumps(asdict(payload)),
                                      rec.report_text(discr))
        except Exception as e:
            fixes.append(f"correction pass failed: {e}")
            break
        fixes += rec.auto_fix(payload)
        discr = rec.reconcile(payload)
        attempts += 1
    return payload, discr, fixes
