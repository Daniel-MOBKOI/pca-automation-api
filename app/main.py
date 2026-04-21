from __future__ import annotations

import re
import shutil
import tempfile
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pptx import Presentation

app = FastAPI(title="PCA Automation API", version="6.0.0")

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
        "required_placeholders": {
            "{{CAMPAIGN_NAME}}",
            "{{CAMPAIGN_PERIOD}}",
            "{{DELIVERED_IMPRESSIONS}}",
            "{{PERFORMANCE_CTR}}",
            "{{PERFORMANCE_ENGAGEMENT_RATE}}",
            "{{PERFORMANCE_VCR}}",
            "{{PERFORMANCE_ON_SCREEN}}",
        },
        "sheets": [],
    }

    if not rules_path.exists():
        return config

    wb = openpyxl.load_workbook(rules_path, data_only=True)
    config["sheets"] = wb.sheetnames
    return config


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
# Tokens
# -------------------------
KPI_TOKENS = {
    "impressions": [["impressions"]],
    "ctr": [["ctr"], ["click", "through"]],
    "engagement_rate": [["engagement", "rate"], ["engagement", "%"], ["er"]],
    "vcr": [["video", "completion", "rate"], ["vcr"], ["completion"]],
    "on_screen": [["mobkoi", "on", "screen"], ["on", "screen"], ["viewability"]],
    "spend": [["actual", "spend"], ["spend"], ["budget"]],
}

DELIVERY_TOKENS = {
    "io_overall_impressions": [["sold", "paid", "units"]],
    "delivered_overall_av_units": [["delivered", "overall", "av", "units"], ["delivered", "overall", "av"]],
    "delivery_with_av_percent": [["delivery", "percentage", "incl", "av"], ["delivery", "incl", "av"]],
    "added_value_worth": [["delivered", "av", "amount"], ["av", "amount"], ["amount"]],
}

SITE_TOKENS = {
    "ctr": KPI_TOKENS["ctr"],
    "er": KPI_TOKENS["engagement_rate"],
    "vcr": KPI_TOKENS["vcr"],
}


def score_kpi_table(hmap: Dict[str, int]) -> int:
    return sum(
        1 for key in ["impressions", "ctr", "engagement_rate", "vcr", "on_screen", "spend"]
        if find_column(hmap, KPI_TOKENS[key]) is not None
    )


def score_delivery_table(hmap: Dict[str, int]) -> int:
    return sum(
        1 for key in ["io_overall_impressions", "delivered_overall_av_units", "delivery_with_av_percent", "added_value_worth"]
        if find_column(hmap, DELIVERY_TOKENS[key]) is not None
    )


# -------------------------
# Business logic
# -------------------------
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
    match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", filename)
    if not match:
        return None, None

    end = datetime(int(match.group(3)), int(match.group(2)), int(match.group(1))) - timedelta(days=1)
    start = end - timedelta(days=6)
    return start, end


def detect_placeholders_in_ppt(template_path: Path) -> set[str]:
    prs = Presentation(str(template_path))
    found = set()
    pattern = re.compile(r"\{\{[A-Z0-9_]+\}\}")

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


# -------------------------
# Normalisation / validation
# -------------------------
def normalize_mapped_values(mapped: Dict[str, str]) -> Dict[str, str]:
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


def validate_export_ready(mapped: Dict[str, str]) -> Tuple[bool, List[str]]:
    required = [
        "{{CAMPAIGN_NAME}}",
        "{{CAMPAIGN_PERIOD}}",
        "{{DELIVERED_IMPRESSIONS}}",
        "{{PERFORMANCE_CTR}}",
        "{{PERFORMANCE_ENGAGEMENT_RATE}}",
        "{{PERFORMANCE_VCR}}",
        "{{PERFORMANCE_ON_SCREEN}}",
    ]
    missing = [k for k in required if mapped.get(k, "N/A") == "N/A"]
    return len(missing) == 0, missing


# -------------------------
# Safe getters
# -------------------------
def safe_value(ws, row: int, col: Optional[int]) -> Any:
    if not row or not col or row < 1 or col < 1:
        return None
    return ws.cell(row, col).value


def unique_join(values: List[str]) -> str:
    out = []
    for v in values:
        if v and v not in out:
            out.append(v)
    return ", ".join(out) if out else "N/A"


# -------------------------
# Core mapping engine
# -------------------------
def map_eoc(eoc_path: Path, template_path: Optional[Path], rules_config: Dict[str, Any]) -> Tuple[Dict[str, str], Dict[str, Any]]:
    wb = openpyxl.load_workbook(eoc_path, data_only=True)
    preferred_sheet = "Consolidated Report"
    ws = wb[preferred_sheet] if preferred_sheet in wb.sheetnames else wb[wb.sheetnames[0]]

    labels = row_labels(ws)
    campaign_rows = labels.get("campaign", [])
    geo_rows = labels.get("geo", []) or labels.get("market", []) or labels.get("country", []) or labels.get("region", [])
    format_rows = labels.get("format", [])
    site_rows = labels.get("site", [])
    date_rows = labels.get("date", [])

    logs: Dict[str, Any] = {
        "sheet_used": ws.title,
        "campaign_rows": campaign_rows,
        "geo_rows": geo_rows,
        "format_rows": format_rows,
        "site_rows": site_rows,
        "date_rows": date_rows,
        "warnings": [],
        "rules_sheets": rules_config.get("sheets", []),
    }

    if not campaign_rows:
        raise ValueError("No Campaign table found in EOC report")

    mapped: Dict[str, str] = {}

    # Detect KPI / Delivery tables
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

    if kpi_header == delivery_header and len(campaign_rows) > 1:
        for row in campaign_rows:
            if row != kpi_header and score_delivery_table(header_map(ws, row)) > 0:
                delivery_header = row
                break

    logs["kpi_header"] = kpi_header
    logs["delivery_header"] = delivery_header

    # KPI extraction
    campaign_name_raw = None
    if kpi_header:
        hmap = header_map(ws, kpi_header)
        data_row = kpi_header + 1

        campaign_name_raw = safe_value(ws, data_row, 2)
        mapped["{{CAMPAIGN_NAME}}"] = str(campaign_name_raw or "N/A")
        mapped["{{DELIVERED_IMPRESSIONS}}"] = fmt_int(safe_value(ws, data_row, find_column(hmap, KPI_TOKENS["impressions"])))
        mapped["{{PERFORMANCE_CTR}}"] = fmt_pct(safe_value(ws, data_row, find_column(hmap, KPI_TOKENS["ctr"])), 2)
        mapped["{{PERFORMANCE_ENGAGEMENT_RATE}}"] = fmt_pct(safe_value(ws, data_row, find_column(hmap, KPI_TOKENS["engagement_rate"])), 2)
        mapped["{{PERFORMANCE_VCR}}"] = fmt_pct(safe_value(ws, data_row, find_column(hmap, KPI_TOKENS["vcr"])), 1)
        mapped["{{PERFORMANCE_ON_SCREEN}}"] = fmt_pct(safe_value(ws, data_row, find_column(hmap, KPI_TOKENS["on_screen"])), 1)
        mapped["{{CAMPAIGN_BUDGET}}"] = fmt_cur(safe_value(ws, data_row, find_column(hmap, KPI_TOKENS["spend"])))
    else:
        logs["warnings"].append("No KPI table found")
        for key in [
            "{{CAMPAIGN_NAME}}",
            "{{DELIVERED_IMPRESSIONS}}",
            "{{PERFORMANCE_CTR}}",
            "{{PERFORMANCE_ENGAGEMENT_RATE}}",
            "{{PERFORMANCE_VCR}}",
            "{{PERFORMANCE_ON_SCREEN}}",
            "{{CAMPAIGN_BUDGET}}",
        ]:
            mapped[key] = "N/A"

    # Delivery extraction
    if delivery_header:
        hmap = header_map(ws, delivery_header)
        data_row = delivery_header + 1

        io_val = safe_value(ws, data_row, find_column(hmap, DELIVERY_TOKENS["io_overall_impressions"]))
        av_val = safe_value(ws, data_row, find_column(hmap, DELIVERY_TOKENS["delivered_overall_av_units"]))
        pct_val = safe_value(ws, data_row, find_column(hmap, DELIVERY_TOKENS["delivery_with_av_percent"]))
        worth_val = safe_value(ws, data_row, find_column(hmap, DELIVERY_TOKENS["added_value_worth"]))

        mapped["{{IO_OVERALL_IMPRESSIONS}}"] = fmt_int(io_val)
        mapped["{{DELIVERED_OVERALL_AV_UNITS}}"] = fmt_int(av_val, absolute=True)
        mapped["{{ADDED_VALUE_IMPRESSIONS}}"] = fmt_int(av_val, absolute=True)
        mapped["{{DELIVERY_WITH_AV_PERCENT}}"] = fmt_pct(pct_val, 1)
        mapped["{{ADDED_VALUE_WORTH}}"] = fmt_cur(worth_val)
    else:
        for key in [
            "{{IO_OVERALL_IMPRESSIONS}}",
            "{{DELIVERED_OVERALL_AV_UNITS}}",
            "{{ADDED_VALUE_IMPRESSIONS}}",
            "{{DELIVERY_WITH_AV_PERCENT}}",
            "{{ADDED_VALUE_WORTH}}",
        ]:
            mapped[key] = "N/A"

    # Dates
    start_date, end_date = detect_date_range_from_table(ws, date_rows[0] if date_rows else None)
    if start_date is None or end_date is None:
        start_date, end_date = fallback_weekly_dates(eoc_path.name)
        if start_date and end_date:
            logs["warnings"].append("Used weekly filename fallback for dates")

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
        values = []
        for r in collect_rows(ws, format_rows[0] + 1):
            val = ws.cell(r, 2).value
            if val not in (None, ""):
                values.append(str(val).strip())
        mapped["{{CAMPAIGN_FORMATS}}"] = unique_join(values)
    else:
        mapped["{{CAMPAIGN_FORMATS}}"] = "N/A"

    # Markets
    if geo_rows:
        values = []
        for r in collect_rows(ws, geo_rows[0] + 1):
            val = ws.cell(r, 2).value
            if val not in (None, ""):
                values.append(str(val).strip())
        mapped["{{CAMPAIGN_MARKETS}}"] = unique_join(values)
    else:
        mapped["{{CAMPAIGN_MARKETS}}"] = "N/A"

    # Client
    mapped["{{CLIENT_NAME}}"] = detect_brand(eoc_path.name, str(campaign_name_raw) if campaign_name_raw else None)

    # Top titles
    max_rank = 5
    if template_path and template_path.exists():
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

        ctr_col = find_column(hmap, SITE_TOKENS["ctr"])
        er_col = find_column(hmap, SITE_TOKENS["er"])
        vcr_col = find_column(hmap, SITE_TOKENS["vcr"])

        rows = collect_rows(ws, site_header + 1)
        entries = []
        for r in rows:
            name = ws.cell(r, 2).value
            if name in (None, ""):
                continue
            entries.append({
                "name": str(name),
                "ctr": safe_value(ws, r, ctr_col),
                "er": safe_value(ws, r, er_col),
                "vcr": safe_value(ws, r, vcr_col),
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

    mapped = normalize_mapped_values(mapped)

    required = set(rules_config.get("required_placeholders", set()))
    unresolved_required = [placeholder for placeholder in required if mapped.get(placeholder, "N/A") == "N/A"]
    logs["unresolved_required"] = sorted(unresolved_required)

    return mapped, logs


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
    return {
        "status": "ok",
        "rules_workbook_found": RULES_WORKBOOK.exists(),
        "default_template_found": DEFAULT_TEMPLATE.exists(),
    }


@app.post("/validate-eoc")
async def validate_eoc(eoc_file: UploadFile = File(...)):
    try:
        if not RULES_WORKBOOK.exists():
            raise HTTPException(status_code=500, detail="Rules workbook not found in assets/")
        if not DEFAULT_TEMPLATE.exists():
            raise HTTPException(status_code=500, detail="Default template not found in assets/")

        template_path = DEFAULT_TEMPLATE

        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            eoc_path = tmp / eoc_file.filename

            with eoc_path.open("wb") as f:
                shutil.copyfileobj(eoc_file.file, f)

            rules_config = load_rules_config(RULES_WORKBOOK)
            mapped, logs = map_eoc(eoc_path, template_path, rules_config)
            return JSONResponse(content={"mapped_values": mapped, "logs": logs})

    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/generate-exec-summary")
async def generate_exec_summary(eoc_file: UploadFile = File(...)):
    try:
        if not RULES_WORKBOOK.exists():
            raise HTTPException(status_code=500, detail="Rules workbook not found in assets/")
        if not DEFAULT_TEMPLATE.exists():
            raise HTTPException(status_code=500, detail="Default template not found in assets/")

        template_path = DEFAULT_TEMPLATE

        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            eoc_path = tmp / eoc_file.filename

            with eoc_path.open("wb") as f:
                shutil.copyfileobj(eoc_file.file, f)

            rules_config = load_rules_config(RULES_WORKBOOK)
            mapped, logs = map_eoc(eoc_path, template_path, rules_config)

            ok, missing = validate_export_ready(mapped)
            if not ok:
                raise HTTPException(
                    status_code=422,
                    detail={
                        "message": "Unable to complete - please check file structure.",
                        "missing": missing,
                        "mapped_values": mapped,
                        "logs": logs,
                    },
                )

            out_path = tmp / "output.pptx"
            replace_placeholders(template_path, out_path, mapped)

            final = OUTPUT_DIR / "Exec_Summary_Output.pptx"
            shutil.copy2(out_path, final)

            return FileResponse(
                path=str(final),
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                filename="Exec_Summary_Output.pptx",
            )

    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )
