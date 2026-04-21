from __future__ import annotations

import re
import shutil
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel
from pptx import Presentation

app = FastAPI(title="PCA Automation API", version="2.0.0")

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# -------------------------
# Formatting helpers
# -------------------------
def norm(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()

def is_num(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)

def fmt_pct(value: Any, decimals: int = 2) -> str:
    if not is_num(value):
        return "N/A"
    return f"{value * 100:.{decimals}f}%"

def fmt_int(value: Any, absolute: bool = False) -> str:
    if not is_num(value):
        return "N/A"
    if absolute:
        value = abs(value)
    return f"{int(round(value)):,}"

def fmt_cur(value: Any, absolute: bool = True) -> str:
    if not is_num(value):
        return "N/A"
    if absolute:
        value = abs(value)
    return f"€{value:,.2f}"

# -------------------------
# Table helpers
# -------------------------
def collect_rows(ws, start_row: int) -> List[int]:
    rows = []
    r = start_row
    while r <= ws.max_row and ws.cell(r, 2).value not in (None, ""):
        rows.append(r)
        r += 1
    return rows

def row_labels(ws) -> Dict[str, List[int]]:
    labels: Dict[str, List[int]] = {}
    for r in range(1, ws.max_row + 1):
        label = norm(ws.cell(r, 2).value)
        if label:
            labels.setdefault(label, []).append(r)
    return labels

def header_map(ws, row: int) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for c in range(2, ws.max_column + 1):
        val = ws.cell(row, c).value
        if val not in (None, ""):
            out[norm(val)] = c
    return out

def find_column(hmap: Dict[str, int], tokens: List[List[str]]) -> Optional[int]:
    for group in tokens:
        for h, col in hmap.items():
            if all(t in h for t in group):
                return col
    return None

# -------------------------
# Token rules
# -------------------------
KPI = {
    "impressions": [["impressions"]],
    "ctr": [["ctr"]],
    "er": [["engagement", "rate"]],
    "vcr": [["vcr"], ["completion"]],
    "view": [["viewability"], ["on", "screen"]],
    "spend": [["spend"]],
}

DELIVERY = {
    "io": [["sold"]],
    "av": [["av"]],
    "pct": [["delivery"]],
}

# -------------------------
# Core mapping
# -------------------------
def map_eoc(eoc_path: Path) -> Dict[str, str]:
    wb = openpyxl.load_workbook(eoc_path, data_only=True)
    ws = wb.active

    labels = row_labels(ws)
    campaign_row = labels.get("campaign", [None])[0]

    mapped = {}

    if campaign_row:
        hmap = header_map(ws, campaign_row)
        r = campaign_row + 1

        mapped["{{CAMPAIGN_NAME}}"] = str(ws.cell(r, 2).value or "N/A")
        mapped["{{DELIVERED_IMPRESSIONS}}"] = fmt_int(ws.cell(r, find_column(hmap, KPI["impressions"]) or 0).value)
        mapped["{{PERFORMANCE_CTR}}"] = fmt_pct(ws.cell(r, find_column(hmap, KPI["ctr"]) or 0).value)
        mapped["{{PERFORMANCE_ENGAGEMENT_RATE}}"] = fmt_pct(ws.cell(r, find_column(hmap, KPI["er"]) or 0).value)
        mapped["{{PERFORMANCE_VCR}}"] = fmt_pct(ws.cell(r, find_column(hmap, KPI["vcr"]) or 0).value)
        mapped["{{PERFORMANCE_ON_SCREEN}}"] = fmt_pct(ws.cell(r, find_column(hmap, KPI["view"]) or 0).value)
        mapped["{{CAMPAIGN_BUDGET}}"] = fmt_cur(ws.cell(r, find_column(hmap, KPI["spend"]) or 0).value)
    else:
        for k in [
            "{{CAMPAIGN_NAME}}",
            "{{DELIVERED_IMPRESSIONS}}",
            "{{PERFORMANCE_CTR}}",
            "{{PERFORMANCE_ENGAGEMENT_RATE}}",
            "{{PERFORMANCE_VCR}}",
            "{{PERFORMANCE_ON_SCREEN}}",
            "{{CAMPAIGN_BUDGET}}",
        ]:
            mapped[k] = "N/A"

    # Simple dates fallback
    mapped["{{CAMPAIGN_PERIOD}}"] = "Q1 2026"

    return mapped

# -------------------------
# FIXED REPLACEMENT FUNCTION
# -------------------------
def replace_placeholders(template_path: Path, out_path: Path, repl: Dict[str, str]):
    prs = Presentation(str(template_path))

    def replace_in_text_frame(tf):
        for p in tf.paragraphs:
            for run in p.runs:
                text = run.text
                if not text:
                    continue

                new_text = text
                for placeholder, value in repl.items():
                    if placeholder in new_text:
                        new_text = new_text.replace(placeholder, str(value))

                run.text = new_text

    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                replace_in_text_frame(shape.text_frame)

            if hasattr(shape, "table"):
                try:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            replace_in_text_frame(cell.text_frame)
                except:
                    pass

    prs.save(str(out_path))

# -------------------------
# API
# -------------------------
@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/generate-exec-summary")
async def generate_exec_summary(
    eoc_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        eoc_path = tmp / eoc_file.filename
        template_path = tmp / template_file.filename

        with eoc_path.open("wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        with template_path.open("wb") as f:
            shutil.copyfileobj(template_file.file, f)

        mapped = map_eoc(eoc_path)

        out_path = tmp / "output.pptx"
        replace_placeholders(template_path, out_path, mapped)

        return FileResponse(
            path=str(out_path),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename="Exec_Summary.pptx",
        )
