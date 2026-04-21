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
# Table detection helpers
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
        value = ws.cell(row, c).value
        if value not in (None, ""):
            out[norm(value)] = c
    return out

def find_column_by_tokens(hmap: Dict[str, int], token_groups: List[List[str]]) -> Optional[int]:
    for tokens in token_groups:
        for header, col in hmap.items():
            if all(token in header for token in tokens):
                return col
    return None

# -------------------------
# Rule dictionaries
# -------------------------
KPI_TOKENS = {
    "impressions": [["impressions"]],
    "ctr": [["ctr"], ["click", "through"]],
    "engagement_rate": [["engagement", "rate"], ["er"]],
    "vcr": [["video", "completion", "rate"], ["vcr"], ["completed", "view", "rate"]],
    "on_screen": [["mobkoi", "on", "screen"], ["on", "screen"], ["viewability"]],
    "spend": [["actual", "spend"], ["spend"], ["total", "spend"], ["budget"]],
}

DELIVERY_TOKENS = {
    "io_overall_impressions": [["sold", "paid", "units"]],
    "delivered_overall_av_units": [["delivered", "overall", "av", "units"], ["delivered", "overall", "av"]],
    "delivery_with_av_percent": [["delivery", "percentage", "incl", "av"], ["delivery", "incl", "av"]],
    "added_value_worth": [["delivered", "av", "amount"], ["av", "amount"]],
}

SITE_TOKENS = {
    "ctr": KPI_TOKENS["ctr"],
    "er": KPI_TOKENS["engagement_rate"],
    "vcr": KPI_TOKENS["vcr"],
}

def score_kpi_table(hmap: Dict[str, int]) -> int:
    score = 0
    for groups in KPI_TOKENS.values():
        if find_column_by_tokens(hmap, groups) is not None:
            score += 1
    return score

def score_delivery_table(hmap: Dict[str, int]) -> int:
    score = 0
    for groups in DELIVERY_TOKENS.values():
        if find_column_by_tokens(hmap, groups) is not None:
            score += 1
    return score

# -------------------------
# Business logic helpers
# -------------------------
def detect_brand(filename: str, campaign_name: Optional[str] = None) -> str:
    stem = Path(filename).stem
    if " - " in stem:
        first = stem.split(" - ")[0].strip()
        if first:
            return first
    if campaign_name:
        return str(campaign_name).split()[0]
    return "N/A"

def detect_date_range_from_table(ws, header_row: Optional[int]) -> Tuple[Optional[datetime], Optional[datetime]]:
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

def fallback_weekly_dates(filename: str) -> Tuple[Optional[datetime], Optional[datetime]]:
    match = re.search(r"(\\d{2})\\.(\\d{2})\\.(\\d{4})", filename)
    if not match:
        return None, None
    # Weekly reports: filename date is the day after the weekly range ended in our Church's case
    end = datetime(int(match.group(3)), int(match.group(2)), int(match.group(1))) - timedelta(days=1)
    start = end - timedelta(days=6)
    return start, end

def detect_placeholders_in_ppt(template_path: Path) -> set[str]:
    prs = Presentation(str(template_path))
    found = set()
    pattern = re.compile(r"\\{\\{[A-Z0-9_]+\\}\\}")

    def scan_text_frame(tf):
        for p in tf.paragraphs:
            for run in p.runs:
                found.update(pattern.findall(run.text or ""))

    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                scan_text_frame(shape.text_frame)
            if hasattr(shape, "table"):
                try:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            scan_text_frame(cell.text_frame)
                except Exception:
                    pass
    return found

def replace_placeholders_in_ppt(template_path: Path, out_path: Path, repl: Dict[str, str]) -> None:
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
# Core mapping engine
# -------------------------
def map_eoc(eoc_path: Path, template_path: Path) -> Tuple[Dict[str, str], Dict[str, Any]]:
    wb = openpyxl.load_workbook(eoc_path, data_only=True)
    ws = wb["Consolidated Report"] if "Consolidated Report" in wb.sheetnames else wb.active

    labels = row_labels(ws)
    campaign_rows = labels.get("campaign", [])
    geo_rows = labels.get("geo", []) or labels.get("market", []) or labels.get("country", []) or labels.get("region", [])
    format_rows = labels.get("format", [])
    site_rows = labels.get("site", [])
    date_rows = labels.get("date", [])

    logs: Dict[str, Any] = {
        "campaign_rows": campaign_rows,
        "geo_rows": geo_rows,
        "format_rows": format_rows,
        "site_rows": site_rows,
        "date_rows": date_rows,
        "warnings": [],
    }

    # Detect KPI + delivery tables
    kpi_header = None
    delivery_header = None
    best_kpi_score = -1
    best_delivery_score = -1

    for row in campaign_rows:
        hmap = header_map(ws, row)
        kpi_score = score_kpi_table(hmap)
        delivery_score = score_delivery_table(hmap)

        if kpi_score > best_kpi_score:
            best_kpi_score = kpi_score
            kpi_header = row

        if delivery_score > best_delivery_score:
            best_delivery_score = delivery_score
            delivery_header = row

    # If same table selected for both and there are multiple campaign tables, split them
    if kpi_header == delivery_header and len(campaign_rows) > 1:
        for row in campaign_rows:
            if row != kpi_header and score_delivery_table(header_map(ws, row)) > 0:
                delivery_header = row
                break

    logs["kpi_header"] = kpi_header
    logs["delivery_header"] = delivery_header

    mapped: Dict[str, str] = {}

    # KPI table extraction
    campaign_name_raw = None
    if kpi_header:
        r = kpi_header + 1
        hmap = header_map(ws, kpi_header)
        campaign_name_raw = ws.cell(r, 2).value

        mapped["{{CAMPAIGN_NAME}}"] = str(campaign_name_raw or "N/A")
        mapped["{{DELIVERED_IMPRESSIONS}}"] = fmt_int(ws.cell(r, find_column_by_tokens(hmap, KPI_TOKENS["impressions"]) or 0).value)
        mapped["{{PERFORMANCE_CTR}}"] = fmt_pct(ws.cell(r, find_column_by_tokens(hmap, KPI_TOKENS["ctr"]) or 0).value, 2)
        mapped["{{PERFORMANCE_ENGAGEMENT_RATE}}"] = fmt_pct(ws.cell(r, find_column_by_tokens(hmap, KPI_TOKENS["engagement_rate"]) or 0).value, 2)
        mapped["{{PERFORMANCE_VCR}}"] = fmt_pct(ws.cell(r, find_column_by_tokens(hmap, KPI_TOKENS["vcr"]) or 0).value, 1)
        mapped["{{PERFORMANCE_ON_SCREEN}}"] = fmt_pct(ws.cell(r, find_column_by_tokens(hmap, KPI_TOKENS["on_screen"]) or 0).value, 1)
        mapped["{{CAMPAIGN_BUDGET}}"] = fmt_cur(ws.cell(r, find_column_by_tokens(hmap, KPI_TOKENS["spend"]) or 0).value)
    else:
        logs["warnings"].append("No KPI table found")
        for key in [
            "{{CAMPAIGN_NAME}}", "{{DELIVERED_IMPRESSIONS}}", "{{PERFORMANCE_CTR}}",
            "{{PERFORMANCE_ENGAGEMENT_RATE}}", "{{PERFORMANCE_VCR}}",
            "{{PERFORMANCE_ON_SCREEN}}", "{{CAMPAIGN_BUDGET}}"
        ]:
            mapped[key] = "N/A"

    # Delivery table extraction
    if delivery_header:
        r = delivery_header + 1
        hmap = header_map(ws, delivery_header)

        io_val = ws.cell(r, find_column_by_tokens(hmap, DELIVERY_TOKENS["io_overall_impressions"]) or 0).value
        av_val = ws.cell(r, find_column_by_tokens(hmap, DELIVERY_TOKENS["delivered_overall_av_units"]) or 0).value
        deliv_pct_val = ws.cell(r, find_column_by_tokens(hmap, DELIVERY_TOKENS["delivery_with_av_percent"]) or 0).value
        av_worth_val = ws.cell(r, find_column_by_tokens(hmap, DELIVERY_TOKENS["added_value_worth"]) or 0).value

        mapped["{{IO_OVERALL_IMPRESSIONS}}"] = fmt_int(io_val)
        mapped["{{ADDED_VALUE_IMPRESSIONS}}"] = fmt_int(av_val, absolute=True)
        mapped["{{DELIVERED_OVERALL_AV_UNITS}}"] = fmt_int(av_val, absolute=True)
        mapped["{{DELIVERY_WITH_AV_PERCENT}}"] = fmt_pct(deliv_pct_val, 1)
        mapped["{{ADDED_VALUE_WORTH}}"] = fmt_cur(av_worth_val)
    else:
        for key in [
            "{{IO_OVERALL_IMPRESSIONS}}", "{{ADDED_VALUE_IMPRESSIONS}}",
            "{{DELIVERED_OVERALL_AV_UNITS}}", "{{DELIVERY_WITH_AV_PERCENT}}",
            "{{ADDED_VALUE_WORTH}}"
        ]:
            mapped[key] = "N/A"

    # Dates
    start_date, end_date = detect_date_range_from_table(ws, date_rows[0] if date_rows else None)
    if start_date is None or end_date is None:
        start_date, end_date = fallback_weekly_dates(eoc_path.name)
        if start_date and end_date:
            logs["warnings"].append("Used filename fallback for dates")

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
        values = [str(ws.cell(r, 2).value) for r in collect_rows(ws, format_rows[0] + 1) if ws.cell(r, 2).value not in (None, "")]
        mapped["{{CAMPAIGN_FORMATS}}"] = ", ".join(values) if values else "N/A"
    else:
        mapped["{{CAMPAIGN_FORMATS}}"] = "N/A"

    # Markets
    if geo_rows:
        values = [str(ws.cell(r, 2).value) for r in collect_rows(ws, geo_rows[0] + 1) if ws.cell(r, 2).value not in (None, "")]
        mapped["{{CAMPAIGN_MARKETS}}"] = ", ".join(values) if values else "N/A"
    else:
        mapped["{{CAMPAIGN_MARKETS}}"] = "N/A"

    # Brand / client
    mapped["{{CLIENT_NAME}}"] = detect_brand(eoc_path.name, str(campaign_name_raw) if campaign_name_raw else None)

    # Top titles
    placeholders = detect_placeholders_in_ppt(template_path)
    highest_rank = 0
    for metric in ["CTR", "ER", "VCR"]:
        for i in range(1, 6):
            if f"{{{{TOP_TITLES_{metric}_{i}_NAME}}}}" in placeholders:
                highest_rank = max(highest_rank, i)
    max_rank = highest_rank or 5

    if site_rows:
        site_header = site_rows[0]
        hmap = header_map(ws, site_header)

        ctr_col = find_column_by_tokens(hmap, SITE_TOKENS["ctr"])
        er_col = find_column_by_tokens(hmap, SITE_TOKENS["er"])
        vcr_col = find_column_by_tokens(hmap, SITE_TOKENS["vcr"])

        rows = collect_rows(ws, site_header + 1)
        entries = []
        for r in rows:
            entries.append({
                "name": str(ws.cell(r, 2).value),
                "ctr": ws.cell(r, ctr_col).value if ctr_col else None,
                "er": ws.cell(r, er_col).value if er_col else None,
                "vcr": ws.cell(r, vcr_col).value if vcr_col else None,
            })

        for prefix, metric, decimals in [("CTR", "ctr", 2), ("ER", "er", 2), ("VCR", "vcr", 1)]:
            ranked = sorted(entries, key=lambda x: x[metric] or 0, reverse=True)
            for i in range(1, 6):
                name_key = f"{{{{TOP_TITLES_{prefix}_{i}_NAME}}}}"
                value_key = f"{{{{TOP_TITLES_{prefix}_{i}_VALUE}}}}"
                if i <= max_rank and i <= len(ranked) and ranked[i - 1][metric] is not None:
                    mapped[name_key] = ranked[i - 1]["name"]
                    mapped[value_key] = fmt_pct(ranked[i - 1][metric], decimals)
                else:
                    mapped[name_key] = "N/A"
                    mapped[value_key] = "N/A"
    else:
        for prefix in ["CTR", "ER", "VCR"]:
            for i in range(1, 6):
                mapped[f"{{{{TOP_TITLES_{prefix}_{i}_NAME}}}}"] = "N/A"
                mapped[f"{{{{TOP_TITLES_{prefix}_{i}_VALUE}}}}"] = "N/A"

    # Required export gate
    required = [
        "{{CAMPAIGN_NAME}}",
        "{{CAMPAIGN_PERIOD}}",
        "{{DELIVERED_IMPRESSIONS}}",
        "{{PERFORMANCE_CTR}}",
        "{{PERFORMANCE_VCR}}",
        "{{CAMPAIGN_BUDGET}}",
    ]
    logs["unresolved_required"] = [key for key in required if mapped.get(key, "N/A") == "N/A"]

    return mapped, logs

# -------------------------
# API models
# -------------------------
class ValidateResponse(BaseModel):
    mapped_values: Dict[str, Any]
    logs: Dict[str, Any]

# -------------------------
# API routes
# -------------------------
@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/validate-eoc", response_model=ValidateResponse)
async def validate_eoc(
    eoc_file: UploadFile = File(...),
    template_file: UploadFile | None = File(None),
):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        eoc_path = tmp / eoc_file.filename
        with eoc_path.open("wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        if template_file:
            template_path = tmp / template_file.filename
            with template_path.open("wb") as f:
                shutil.copyfileobj(template_file.file, f)
        else:
            raise HTTPException(status_code=400, detail="template_file is required for validation")

        mapped, logs = map_eoc(eoc_path, template_path)
        return ValidateResponse(mapped_values=mapped, logs=logs)

@app.post("/generate-exec-summary")
async def generate_exec_summary(
    eoc_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        eoc_path = tmp / eoc_file.filename
        template_path = tmp / template_file.filename

        with eoc_path.open("wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        with template_path.open("wb") as f:
            shutil.copyfileobj(template_file.file, f)

        mapped, logs = map_eoc(eoc_path, template_path)

        if logs["unresolved_required"]:
            raise HTTPException(
                status_code=400,
                detail={
                    "message": "Unable to complete - please check file structure.",
                    "logs": logs,
                },
            )

        out_path = tmp / f"{Path(eoc_file.filename).stem}_Exec_Summary.pptx"
        replace_placeholders_in_ppt(template_path, out_path, mapped)

        final_path = OUTPUT_DIR / out_path.name
        shutil.copy2(out_path, final_path)

        return FileResponse(
            path=str(final_path),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=out_path.name,
        )                        run.text = text

    prs.save(output_path)


# -------------------------
# VALIDATE ENDPOINT
# -------------------------
@app.post("/validate-eoc")
async def validate_eoc(eoc_file: UploadFile = File(...)):
    with tempfile.TemporaryDirectory() as tmp:
        file_path = os.path.join(tmp, eoc_file.filename)

        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(eoc_file.file, buffer)

        data = extract_basic_data(file_path)

        return JSONResponse(content=data)


# -------------------------
# GENERATE PPT
# -------------------------
@app.post("/generate-exec-summary")
async def generate_exec_summary(
    eoc_file: UploadFile = File(...),
    template_file: UploadFile = File(...)
):
    with tempfile.TemporaryDirectory() as tmp:

        eoc_path = os.path.join(tmp, eoc_file.filename)
        template_path = os.path.join(tmp, template_file.filename)

        with open(eoc_path, "wb") as buffer:
            shutil.copyfileobj(eoc_file.file, buffer)

        with open(template_path, "wb") as buffer:
            shutil.copyfileobj(template_file.file, buffer)

        data = extract_basic_data(eoc_path)

        output_file = os.path.join(tmp, "output.pptx")

        replace_text(template_path, output_file, data)

        final_path = OUTPUT_DIR / "Exec_Summary_Output.pptx"
        shutil.copy(output_file, final_path)

        return FileResponse(
            final_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename="Exec_Summary.pptx"
        )
