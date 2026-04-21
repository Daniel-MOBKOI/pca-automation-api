from __future__ import annotations

import re
import shutil
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pptx import Presentation

app = FastAPI(title="PCA Automation API", version="2.2.0")

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

ASSETS_DIR = BASE_DIR / "assets"
DEFAULT_TEMPLATE = ASSETS_DIR / "Executive Summary_PCA_One Pager_MASTER.pptx"

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
# Workbook helpers
# -------------------------
def row_labels(ws) -> Dict[str, List[int]]:
    labels: Dict[str, List[int]] = {}
    for r in range(1, ws.max_row + 1):
        label = norm(ws.cell(r, 2).value)
        if label:
            labels.setdefault(label, []).append(r)
    return labels

def collect_rows(ws, start_row: int) -> List[int]:
    rows = []
    r = start_row
    while r <= ws.max_row and ws.cell(r, 2).value not in (None, ""):
        rows.append(r)
        r += 1
    return rows

def header_map(ws, row: int) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for c in range(2, ws.max_column + 1):
        value = ws.cell(row, c).value
        if value not in (None, ""):
            out[norm(value)] = c
    return out

def find_column(hmap: Dict[str, int], token_groups: List[List[str]]) -> Optional[int]:
    for tokens in token_groups:
        for header, col in hmap.items():
            if all(token in header for token in tokens):
                return col
    return None

# -------------------------
# Detection rules
# -------------------------
KPI_TOKENS = {
    "impressions": [["impressions"]],
    "ctr": [["ctr"]],
    "engagement_rate": [["engagement", "rate"], ["er"]],
    "vcr": [["video", "completion", "rate"], ["vcr"], ["completion"]],
    "on_screen": [["mobkoi", "on", "screen"], ["on", "screen"], ["viewability"]],
    "spend": [["actual", "spend"], ["spend"], ["total", "spend"], ["budget"]],
}

def detect_brand(filename: str, campaign_name: Optional[str] = None) -> str:
    stem = Path(filename).stem
    if " - " in stem:
        first = stem.split(" - ")[0].strip()
        if first:
            return first
    if campaign_name:
        words = str(campaign_name).split()
        if words:
            return words[0]
    return "N/A"

def detect_date_range_from_table(ws, header_row: Optional[int]) -> tuple[Optional[datetime], Optional[datetime]]:
    if not header_row:
        return None, None

    dates = []
    for r in collect_rows(ws, header_row + 1):
        value = ws.cell(r, 2).value
        if isinstance(value, datetime):
            dates.append(value)

    if not dates:
        return None, None

    return min(dates), max(dates)

def fallback_weekly_dates(filename: str) -> tuple[Optional[datetime], Optional[datetime]]:
    match = re.search(r"(\\d{2})\\.(\\d{2})\\.(\\d{4})", filename)
    if not match:
        return None, None

    end = datetime(int(match.group(3)), int(match.group(2)), int(match.group(1))) - timedelta(days=1)
    start = end - timedelta(days=6)
    return start, end

# -------------------------
# Core mapping - Phase 1
# -------------------------
def map_eoc_phase1(eoc_path: Path) -> Dict[str, str]:
    wb = openpyxl.load_workbook(eoc_path, data_only=True)
    ws = wb["Consolidated Report"] if "Consolidated Report" in wb.sheetnames else wb.active

    labels = row_labels(ws)
    campaign_rows = labels.get("campaign", [])
    geo_rows = labels.get("geo", []) or labels.get("market", []) or labels.get("country", []) or labels.get("region", [])
    format_rows = labels.get("format", [])
    date_rows = labels.get("date", [])

    mapped: Dict[str, str] = {}

    # KPI / Campaign table
    campaign_header = None
    best_score = -1

    for row in campaign_rows:
        hmap = header_map(ws, row)
        score = 0
        for key in ["impressions", "ctr", "vcr", "spend"]:
            if find_column(hmap, KPI_TOKENS[key]) is not None:
                score += 1
        if score > best_score:
            best_score = score
            campaign_header = row

    campaign_name_raw = None
    if campaign_header:
        hmap = header_map(ws, campaign_header)
        data_row = campaign_header + 1

        campaign_name_raw = ws.cell(data_row, 2).value
        mapped["{{CAMPAIGN_NAME}}"] = str(campaign_name_raw or "N/A")

        impressions_col = find_column(hmap, KPI_TOKENS["impressions"])
        ctr_col = find_column(hmap, KPI_TOKENS["ctr"])
        er_col = find_column(hmap, KPI_TOKENS["engagement_rate"])
        vcr_col = find_column(hmap, KPI_TOKENS["vcr"])
        on_screen_col = find_column(hmap, KPI_TOKENS["on_screen"])
        spend_col = find_column(hmap, KPI_TOKENS["spend"])

        mapped["{{DELIVERED_IMPRESSIONS}}"] = fmt_int(ws.cell(data_row, impressions_col or 0).value)
        mapped["{{PERFORMANCE_CTR}}"] = fmt_pct(ws.cell(data_row, ctr_col or 0).value, 2)
        mapped["{{PERFORMANCE_ENGAGEMENT_RATE}}"] = fmt_pct(ws.cell(data_row, er_col or 0).value, 2)
        mapped["{{PERFORMANCE_VCR}}"] = fmt_pct(ws.cell(data_row, vcr_col or 0).value, 1)
        mapped["{{PERFORMANCE_ON_SCREEN}}"] = fmt_pct(ws.cell(data_row, on_screen_col or 0).value, 1)
        mapped["{{CAMPAIGN_BUDGET}}"] = fmt_cur(ws.cell(data_row, spend_col or 0).value)
    else:
        mapped["{{CAMPAIGN_NAME}}"] = "N/A"
        mapped["{{DELIVERED_IMPRESSIONS}}"] = "N/A"
        mapped["{{PERFORMANCE_CTR}}"] = "N/A"
        mapped["{{PERFORMANCE_ENGAGEMENT_RATE}}"] = "N/A"
        mapped["{{PERFORMANCE_VCR}}"] = "N/A"
        mapped["{{PERFORMANCE_ON_SCREEN}}"] = "N/A"
        mapped["{{CAMPAIGN_BUDGET}}"] = "N/A"

    # AV / Delivery left as N/A in phase 1
    for key in [
        "{{IO_OVERALL_IMPRESSIONS}}",
        "{{ADDED_VALUE_IMPRESSIONS}}",
        "{{DELIVERED_OVERALL_AV_UNITS}}",
        "{{DELIVERY_WITH_AV_PERCENT}}",
        "{{ADDED_VALUE_WORTH}}",
    ]:
        mapped[key] = "N/A"

    # Dates
    start_date, end_date = detect_date_range_from_table(ws, date_rows[0] if date_rows else None)
    if start_date is None or end_date is None:
        start_date, end_date = fallback_weekly_dates(eoc_path.name)

    if start_date and end_date:
        mapped["{{LIVE_DATES_FULL}}"] = f"{start_date.strftime('%d %B')} - {end_date.strftime('%d %B %Y')}"
        mapped["{{LIVE_DATES_SHORT}}"] = f"{start_date.strftime('%d %b')} - {end_date.strftime('%d %b')}"
        mapped["{{CAMPAIGN_PERIOD}}"] = f"Q{((start_date.month - 1) // 3) + 1} {start_date.year}"
    else:
        mapped["{{LIVE_DATES_FULL}}"] = "N/A"
        mapped["{{LIVE_DATES_SHORT}}"] = "N/A"
        mapped["{{CAMPAIGN_PERIOD}}"] = "N/A"

    # Formats
    if format_rows:
        values = [
            str(ws.cell(r, 2).value)
            for r in collect_rows(ws, format_rows[0] + 1)
            if ws.cell(r, 2).value not in (None, "")
        ]
        mapped["{{CAMPAIGN_FORMATS}}"] = ", ".join(values) if values else "N/A"
    else:
        mapped["{{CAMPAIGN_FORMATS}}"] = "N/A"

    # Markets
    if geo_rows:
        values = [
            str(ws.cell(r, 2).value)
            for r in collect_rows(ws, geo_rows[0] + 1)
            if ws.cell(r, 2).value not in (None, "")
        ]
        mapped["{{CAMPAIGN_MARKETS}}"] = ", ".join(values) if values else "N/A"
    else:
        mapped["{{CAMPAIGN_MARKETS}}"] = "N/A"

    # Brand / Client
    mapped["{{CLIENT_NAME}}"] = detect_brand(eoc_path.name, str(campaign_name_raw) if campaign_name_raw else None)

    # Top titles left as N/A in phase 1
    for prefix in ["CTR", "ER", "VCR"]:
        for i in range(1, 6):
            mapped[f"{{{{TOP_TITLES_{prefix}_{i}_NAME}}}}"] = "N/A"
            mapped[f"{{{{TOP_TITLES_{prefix}_{i}_VALUE}}}}"] = "N/A"

    return mapped

# -------------------------
# PPT replacement
# -------------------------
def replace_placeholders(template_path: Path, out_path: Path, repl: Dict[str, str]) -> None:
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
                except Exception:
                    pass

    prs.save(str(out_path))

# -------------------------
# API
# -------------------------
@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/validate-eoc")
async def validate_eoc(
    eoc_file: UploadFile = File(...),
    template_file: UploadFile | None = File(None),
):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        eoc_path = tmp / eoc_file.filename
        with eoc_path.open("wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        mapped = map_eoc_phase1(eoc_path)
        return JSONResponse(content=mapped)

@app.post("/generate-exec-summary")
async def generate_exec_summary(
    eoc_file: UploadFile = File(...),
    template_file: UploadFile | None = File(None),
):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        eoc_path = tmp / eoc_file.filename
        with eoc_path.open("wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        if template_file:
            template_path = tmp / template_file.filename
            with template_path.open("wb") as f:
                shutil.copyfileobj(template_file.file, f)
        else:
            template_path = DEFAULT_TEMPLATE

        out_path = tmp / "output.pptx"
        mapped = map_eoc_phase1(eoc_path)
        replace_placeholders(template_path, out_path, mapped)

        final_path = OUTPUT_DIR / "Exec_Summary_Output.pptx"
        shutil.copy2(out_path, final_path)

        return FileResponse(
            path=str(final_path),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename="Exec_Summary.pptx",
        )
