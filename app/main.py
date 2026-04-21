from __future__ import annotations

import re
import shutil
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pptx import Presentation

app = FastAPI(title="PCA Automation API", version="4.2.0")

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

ASSETS_DIR = BASE_DIR / "assets"
DEFAULT_TEMPLATE = ASSETS_DIR / "Executive Summary_PCA_One Pager_MASTER.pptx"
RULES_WORKBOOK = ASSETS_DIR / "PCA_GPT_Rules_Master.xlsx"


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
# Rules loader
# -------------------------
def load_rules_config(rules_path: Path) -> Dict[str, Any]:
    config = {
        "required_placeholders": set(),
        "sheets": []
    }

    if not rules_path.exists():
        return config

    wb = openpyxl.load_workbook(rules_path, data_only=True)

    for sheet in wb.sheetnames:
        config["sheets"].append(sheet)

    # core required placeholders
    config["required_placeholders"] = {
        "{{CAMPAIGN_NAME}}",
        "{{CAMPAIGN_PERIOD}}",
        "{{DELIVERED_IMPRESSIONS}}",
        "{{PERFORMANCE_CTR}}",
        "{{PERFORMANCE_ENGAGEMENT_RATE}}",
        "{{PERFORMANCE_VCR}}",
        "{{PERFORMANCE_ON_SCREEN}}",
    }

    return config


# -------------------------
# Helpers
# -------------------------
def row_labels(ws):
    labels = {}
    for r in range(1, ws.max_row + 1):
        val = norm(ws.cell(r, 2).value)
        if val:
            labels.setdefault(val, []).append(r)
    return labels


def collect_rows(ws, start_row):
    rows = []
    r = start_row
    while r <= ws.max_row and ws.cell(r, 2).value not in (None, ""):
        rows.append(r)
        r += 1
    return rows


def header_map(ws, row):
    out = {}
    for c in range(2, ws.max_column + 1):
        val = ws.cell(row, c).value
        if val not in (None, ""):
            out[norm(val)] = c
    return out


def find_column(hmap, token_groups):
    for tokens in token_groups:
        for header, col in hmap.items():
            if all(token in header for token in tokens):
                return col
    return None


# -------------------------
# Tokens
# -------------------------
KPI_TOKENS = {
    "impressions": [["impressions"]],
    "ctr": [["ctr"]],
    "engagement_rate": [["engagement", "rate"], ["er"]],
    "vcr": [["vcr"], ["completion"]],
    "on_screen": [["on", "screen"], ["viewability"]],
    "spend": [["spend"], ["budget"]],
}

DELIVERY_TOKENS = {
    "io": [["sold"]],
    "av": [["av"]],
    "pct": [["delivery"]],
    "worth": [["amount"]],
}


# -------------------------
# Business logic
# -------------------------
def detect_brand(filename, campaign_name=None):
    if " - " in filename:
        return filename.split(" - ")[0]
    if campaign_name:
        return str(campaign_name).split()[0]
    return "N/A"


def detect_dates(ws, date_row):
    if not date_row:
        return None, None

    dates = []
    for r in collect_rows(ws, date_row + 1):
        val = ws.cell(r, 2).value
        if isinstance(val, datetime):
            dates.append(val)

    if not dates:
        return None, None

    return min(dates), max(dates)


def fallback_dates(filename):
    match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", filename)
    if not match:
        return None, None

    end = datetime(int(match.group(3)), int(match.group(2)), int(match.group(1))) - timedelta(days=1)
    start = end - timedelta(days=6)
    return start, end


# -------------------------
# Normalisation
# -------------------------
def normalize_mapped_values(mapped):
    out = dict(mapped)

    if "{{CAMPAIGN_BUDGET}}" in out:
        out["{{ACTUAL_SPEND}}"] = out["{{CAMPAIGN_BUDGET}}"]
        out["{{SPEND}}"] = out["{{CAMPAIGN_BUDGET}}"]

    if "{{DELIVERED_OVERALL_AV_UNITS}}" in out:
        out["{{ADDED_VALUE_IMPRESSIONS}}"] = out["{{DELIVERED_OVERALL_AV_UNITS}}"]

    for k, v in list(out.items()):
        if v is None or v == "":
            out[k] = "N/A"

    return out


# -------------------------
# Validation
# -------------------------
def validate_export_ready(mapped):
    required = [
        "{{CAMPAIGN_NAME}}",
        "{{CAMPAIGN_PERIOD}}",
        "{{DELIVERED_IMPRESSIONS}}",
        "{{PERFORMANCE_CTR}}",
        "{{PERFORMANCE_ENGAGEMENT_RATE}}",
        "{{PERFORMANCE_VCR}}",
        "{{PERFORMANCE_ON_SCREEN}}",
    ]

    missing = [k for k in required if mapped.get(k) == "N/A"]
    return len(missing) == 0, missing


# -------------------------
# Mapping
# -------------------------
def map_eoc(eoc_path, template_path, rules_config):
    wb = openpyxl.load_workbook(eoc_path, data_only=True)
    ws = wb.active

    labels = row_labels(ws)

    campaign_row = labels.get("campaign", [None])[0]
    geo_row = labels.get("geo", [None])[0]
    format_row = labels.get("format", [None])[0]
    date_row = labels.get("date", [None])[0]

    mapped = {}

    # KPI
    hmap = header_map(ws, campaign_row)
    data_row = campaign_row + 1

    mapped["{{CAMPAIGN_NAME}}"] = str(ws.cell(data_row, 2).value)
    mapped["{{DELIVERED_IMPRESSIONS}}"] = fmt_int(ws.cell(data_row, find_column(hmap, KPI_TOKENS["impressions"]) or 0).value)
    mapped["{{PERFORMANCE_CTR}}"] = fmt_pct(ws.cell(data_row, find_column(hmap, KPI_TOKENS["ctr"]) or 0).value)
    mapped["{{PERFORMANCE_ENGAGEMENT_RATE}}"] = fmt_pct(ws.cell(data_row, find_column(hmap, KPI_TOKENS["engagement_rate"]) or 0).value)
    mapped["{{PERFORMANCE_VCR}}"] = fmt_pct(ws.cell(data_row, find_column(hmap, KPI_TOKENS["vcr"]) or 0).value, 1)
    mapped["{{PERFORMANCE_ON_SCREEN}}"] = fmt_pct(ws.cell(data_row, find_column(hmap, KPI_TOKENS["on_screen"]) or 0).value, 1)
    mapped["{{CAMPAIGN_BUDGET}}"] = fmt_cur(ws.cell(data_row, find_column(hmap, KPI_TOKENS["spend"]) or 0).value)

    # Dates
    start, end = detect_dates(ws, date_row)
    if not start:
        start, end = fallback_dates(eoc_path.name)

    if start:
        mapped["{{LIVE_DATES_FULL}}"] = f"{start.strftime('%d %B')} - {end.strftime('%d %B %Y')}"
        mapped["{{LIVE_DATES_SHORT}}"] = f"{start.strftime('%d %b')} - {end.strftime('%d %b')}"
        mapped["{{CAMPAIGN_PERIOD}}"] = f"Q{((start.month - 1)//3)+1} {start.year}"

    # Formats
    if format_row:
        vals = []
        for r in collect_rows(ws, format_row + 1):
            v = ws.cell(r, 2).value
            if v:
                vals.append(str(v))
        mapped["{{CAMPAIGN_FORMATS}}"] = ", ".join(set(vals))

    # Markets
    if geo_row:
        vals = []
        for r in collect_rows(ws, geo_row + 1):
            v = ws.cell(r, 2).value
            if v:
                vals.append(str(v))
        mapped["{{CAMPAIGN_MARKETS}}"] = ", ".join(set(vals))

    mapped["{{CLIENT_NAME}}"] = detect_brand(eoc_path.name)

    mapped = normalize_mapped_values(mapped)

    return mapped, {"status": "ok"}


# -------------------------
# PPT replacement
# -------------------------
def replace_placeholders(template_path, out_path, repl):
    prs = Presentation(str(template_path))

    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        text = run.text
                        for k, v in repl.items():
                            if k in text:
                                text = text.replace(k, str(v))
                        run.text = text

    prs.save(str(out_path))


# -------------------------
# API
# -------------------------
@app.post("/validate-eoc")
async def validate_eoc(eoc_file: UploadFile = File(...)):
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / eoc_file.filename
        with path.open("wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        rules = load_rules_config(RULES_WORKBOOK)
        mapped, logs = map_eoc(path, DEFAULT_TEMPLATE, rules)

        return JSONResponse({"mapped_values": mapped, "logs": logs})


@app.post("/generate-exec-summary")
async def generate_exec_summary(eoc_file: UploadFile = File(...)):
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / eoc_file.filename
        with path.open("wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        rules = load_rules_config(RULES_WORKBOOK)
        mapped, logs = map_eoc(path, DEFAULT_TEMPLATE, rules)

        ok, missing = validate_export_ready(mapped)
        if not ok:
            raise HTTPException(422, detail={"missing": missing})

        out_path = Path(tmp) / "output.pptx"
        replace_placeholders(DEFAULT_TEMPLATE, out_path, mapped)

        final = OUTPUT_DIR / "Exec_Summary_Output.pptx"
        shutil.copy2(out_path, final)

        return FileResponse(final)
