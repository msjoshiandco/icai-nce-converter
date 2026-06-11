# ICAI NCE Financial Statements Conversion Tool

Converts T-format final accounts of a Non-Corporate Entity (proprietorship or
partnership firm) into the ICAI NCE format as a formula-linked Excel workbook.
Built for M S Joshi & Co (FRN 0138082W). FastAPI + Python (openpyxl) + Anthropic Claude.

## Deploy on Render
1. This repo already contains `render.yaml`.
2. Render -> New + -> Blueprint -> pick this repo.
3. Add the secret env var `ANTHROPIC_API_KEY` (your Anthropic key).
4. Deploy. App serves at the Render URL. Health check: `/health`.

## Run locally
    pip install -r requirements.txt
    set ANTHROPIC_API_KEY=sk-ant-...   (optional; only for auto-extract)
    uvicorn main:app --reload
Open http://127.0.0.1:8000 , click "Load sample", then "Generate".

## How it works
Upload both years' statements -> pick constitution -> Claude extracts & classifies
-> review/edit -> generate a tallying, formula-linked .xlsx (Calibri Light 11, no underline).
The build engine needs no API key; only auto-extract does.
