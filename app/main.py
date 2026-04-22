from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, FileResponse
import pandas as pd
import shutil
from pathlib import Path
from datetime import datetime, timedelta
import traceback
import math
import re
import os
from typing import Any, Dict, List, Optional, Tuple
from pptx import Presentation

APP_VERSION = "10.0.1"

app = FastAPI(title="PCA Automation API", version=APP_VERSION)

# -------------------------
# CONFIG
# -------------------------
BASE_DIR = Path("/tmp")
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

RULES_MASTER_PATH = Path(os.getenv("RULES_MASTER_PATH", "PCA_GPT_Rules_Master.xlsx"))

# -------------------------
# HEALTH
# -------------------------
@app.get("/health")
def health():
    return {
        "status": "ok",
        "version": APP_VERSION
    }

# -------------------------
# HELPERS
# -------------------------
def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value)).strip()

def normalize_text(value: Any) -> str:
    return re.sub(r"[^a-z0-9%./()\- ]+", " ", clean_text(value).lower()).strip()

def normalize_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()

def safe_number(value: Any) -> Optional[float]:
    try:
        if value is None:
            return None
        if isinstance(value, str):
            value = (
                value.replace(",", "")
                .replace("£", "")
                .replace("$", "")
                .replace("€", "")
                .replace("%", "")
                .strip()
            )
        num = float(value)
        if math.isnan(num) or math.isinf(num):
            return None
        return num
    except Exception:
        return None

def format_number(value: Optional[float], decimals: int = 0) -> Optional[str]:
    if value is None:
        return None
    if decimals == 0:
        return f"{int(round(value)):,}"
    return f"{value:,.{decimals}f}"

def format_currency(value: Optional[float], symbol: str = "€") -> Optional[str]:
    if value is None:
        return None
    return f"{symbol}{value:,.2f}"

def format_percent(value: Any) -> Optional[str]:
    if isinstance(value, str) and "%" in value:
        return clean_text(value)

    num = safe_number(value)
    if num is None:
        return None

    if num <= 1:
        num = num * 100

    return f"{round(num, 2)}%"

def parse_date_from_any(value: Any) -> Optional[datetime]:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    text = clean_text(value)
    if not text:
        return None

    try:
        dt = pd.to_datetime(text, errors="coerce")
        if pd.notna(dt):
            return dt.to_pydatetime()
    except Exception:
        pass

    match = re.search(r"\d{4}-\d{2}-\d{2}", text)
    if match:
        try:
            return datetime.strptime(match.group(), "%Y-%m-%d")
        except Exception:
            pass

    return None

def format_date_short(start_dt: Optional[datetime], end_dt: Optional[datetime]) -> Optional[str]:
    if start_dt and end_dt:
        return f"{start_dt.strftime('%d %b')} - {end_dt.strftime('%d %b')}"
    return None

def format_date_full(start_dt: Optional[datetime], end_dt: Optional[datetime]) -> Optional[str]:
    if start_dt and end_dt:
        if start_dt.year == end_dt.year:
            return f"{start_dt.strftime('%d %B')} - {end_dt.strftime('%d %B %Y')}"
        return f"{start_dt.strftime('%d %B %Y')} - {end_dt.strftime('%d %B %Y')}"
    return None

def format_quarter_from_date(end_dt: Optional[datetime]) -> Optional[str]:
    if not end_dt:
        return None
    q = ((end_dt.month - 1) // 3) + 1
    return f"Q{q} {end_dt.year}"

def first_non_empty(series: pd.Series) -> Any:
    for value in series:
        if clean_text(value) != "":
            return value
    return None

# -------------------------
# RULES MASTER LOADER
# -------------------------
def load_rules_master() -> Optional[Dict[str, pd.DataFrame]]:
    """
    Optional loader.
    If the Rules Master is not present on Render, continue without it.
    """
    if not RULES_MASTER_PATH.exists():
        return None

    try:
        xls = pd.ExcelFile(RULES_MASTER_PATH, engine="openpyxl")
        out = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(RULES_MASTER_PATH, sheet_name=sheet, engine="openpyxl")
            out[sheet] = df
        return out
    except Exception:
        return None

# -------------------------
# COLUMN MATCHING
# -------------------------
ALIASES = {
    "campaign": ["campaign", "campaign name"],
    "client_brand": ["client", "brand", "advertiser", "brand / client"],
    "impressions": ["impressions", "delivered impressions", "served impressions"],
    "ctr": ["ctr", "click through rate", "click-through rate"],
    "engagement_rate": ["engagement rate", "engagement %", "er", "total er", "overall er"],
    "vcr": ["video completion rate", "vcr", "completed view rate", "video % complete"],
    "on_screen": ["mobkoi on screen", "on screen", "on-screen", "viewability", "viewable rate", "mrc viewability"],
    "spend": ["actual spend", "spend", "total spend", "media spend"],
    "sold_paid_units": ["sold paid units"],
    "delivered_overall_av_units": ["delivered overall av units", "delivered overall av (units)"],
    "delivery_incl_av": ["delivery percentage incl av", "delivery percentage (incl av)"],
    "delivered_av_amount": ["delivered av amount currency", "delivered av amount (currency)"],
    "site": ["site", "publisher", "domain", "environment", "property", "placement"],
    "geo": ["geo", "market", "country", "region"],
    "format": ["format", "creative format", "ad format", "unit type"],
    "date": ["date", "report date", "served date", "live date"],
}

def header_match_score(header: str, aliases: List[str]) -> int:
    norm = normalize_header(header)
    best = 0
    for alias in aliases:
        alias_norm = normalize_header(alias)
        if norm == alias_norm:
            best = max(best, 100)
        elif alias_norm in norm:
            best = max(best, 85)
        elif norm in alias_norm:
            best = max(best, 60)
    return best

def find_col(df: pd.DataFrame, key: str) -> Optional[str]:
    aliases = ALIASES.get(key, [key])
    best_col = None
    best_score = 0
    for col in df.columns:
        score = header_match_score(str(col), aliases)
        if score > best_score:
            best_score = score
            best_col = col
    return best_col if best_score >= 60 else None

# -------------------------
# TABLE DETECTION
# -------------------------
def score_header_row(row_values: List[Any]) -> int:
    keywords = [
        "campaign", "impressions", "ctr", "engagement", "vcr", "completion",
        "spend", "site", "geo", "market", "country", "format", "date"
    ]
    score = 0
    for cell in row_values:
        text = normalize_text(cell)
        for kw in keywords:
            if kw in text:
                score += 1
    return score

def find_header_row(df: pd.DataFrame) -> int:
    best_row = 0
    best_score = -1
    for i in range(min(40, len(df))):
        score = score_header_row(df.iloc[i].tolist())
        if score > best_score:
            best_score = score
            best_row = i
    return best_row

def prepare_sheet_table(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    header_row = find_header_row(df)
    temp = df.iloc[header_row:].copy()
    temp.columns = [clean_text(c) for c in temp.iloc[0]]
    temp = temp[1:].reset_index(drop=True)

    temp = temp.dropna(axis=1, how="all")

    if not temp.empty:
        first_row_norm = [normalize_header(c) for c in temp.columns]
        keep_rows = []
        for _, row in temp.iterrows():
            row_vals = [normalize_header(v) for v in row.tolist()]
            same_count = 0
            for a, b in zip(first_row_norm[:10], row_vals[:10]):
                if a and b and a == b:
                    same_count += 1
            keep_rows.append(same_count < 3)
        temp = temp[pd.Series(keep_rows).values].reset_index(drop=True)

    return temp

def classify_table(df: pd.DataFrame) -> str:
    if find_col(df, "site"):
        return "site"
    if find_col(df, "geo"):
        return "geo"
    if find_col(df, "format"):
        return "format"
    if find_col(df, "date"):
        return "date"

    delivery_hits = sum([
        1 if find_col(df, "sold_paid_units") else 0,
        1 if find_col(df, "delivered_overall_av_units") else 0,
        1 if find_col(df, "delivery_incl_av") else 0,
        1 if find_col(df, "delivered_av_amount") else 0,
    ])
    if delivery_hits >= 2:
        return "campaign_delivery"

    kpi_hits = sum([
        1 if find_col(df, "campaign") else 0,
        1 if find_col(df, "impressions") else 0,
        1 if find_col(df, "ctr") else 0,
        1 if find_col(df, "engagement_rate") else 0,
        1 if find_col(df, "vcr") else 0,
        1 if find_col(df, "on_screen") else 0,
        1 if find_col(df, "spend") else 0,
    ])
    if kpi_hits >= 3:
        return "campaign_kpi_summary"

    return "unknown"

def detect_tables(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    detected = {}

    for _, raw_df in sheets.items():
        if raw_df is None or raw_df.empty:
            continue

        temp = prepare_sheet_table(raw_df)
        if temp.empty:
            continue

        t = classify_table(temp)

        if t != "unknown" and t not in detected:
            detected[t] = temp

    return detected

# -------------------------
# DATE / BRAND EXTRACTION
# -------------------------
def scan_workbook_texts(sheets: Dict[str, pd.DataFrame]) -> List[str]:
    texts = []
    for df in sheets.values():
        if df is None or df.empty:
            continue
        preview = df.head(20)
        for row in preview.values:
            for cell in row:
                text = clean_text(cell)
                if text:
                    texts.append(text)
    return texts

def extract_dates(sheets: Dict[str, pd.DataFrame], detected: Dict[str, pd.DataFrame], filename: str = "") -> Tuple[Optional[datetime], Optional[datetime]]:
    date_df = detected.get("date")
    if date_df is not None:
        date_col = find_col(date_df, "date")
        if date_col:
            vals = [parse_date_from_any(v) for v in date_df[date_col].tolist()]
            vals = [v for v in vals if v]
            if vals:
                return min(vals), max(vals)

    texts = scan_workbook_texts(sheets)
    all_dates = []
    for text in texts:
        dt = parse_date_from_any(text)
        if dt:
            all_dates.append(dt)
    if all_dates:
        return min(all_dates), max(all_dates)

    match = re.search(r"(\d{4}-\d{2}-\d{2})", filename or "")
    if match:
        end_dt = parse_date_from_any(match.group(1))
        if end_dt:
            start_dt = end_dt - timedelta(days=6)
            return start_dt, end_dt

    return None, None

def extract_client_name(sheets: Dict[str, pd.DataFrame], detected: Dict[str, pd.DataFrame], filename: str = "") -> Optional[str]:
    for table_name in ["campaign_kpi_summary", "campaign_delivery"]:
        df = detected.get(table_name)
        if df is not None:
            col = find_col(df, "client_brand")
            if col:
                val = first_non_empty(df[col])
                if clean_text(val):
                    return clean_text(val)

    for text in scan_workbook_texts(sheets):
        low = normalize_text(text)
        if "client:" in low or "brand:" in low:
            parts = text.split(":")
            if len(parts) > 1:
                candidate = clean_text(parts[-1])
                if candidate:
                    return candidate

    if filename:
        base = Path(filename).stem
        if " - " in base:
            return clean_text(base.split(" - ")[0])

    return None

# -------------------------
# CORE MAPPING
# -------------------------
def extract_kpi_value(df: pd.DataFrame, key: str) -> Any:
    col = find_col(df, key)
    if not col:
        return None
    return first_non_empty(df[col])

def join_dimension_values(df: pd.DataFrame, dim_key: str) -> Optional[str]:
    col = find_col(df, dim_key)
    if not col:
        return None

    vals = []
    for value in df[col].tolist():
        txt = clean_text(value)
        if not txt:
            continue
        if txt.lower() in [normalize_header(c) for c in df.columns]:
            continue
        if txt not in vals:
            vals.append(txt)

    if not vals:
        return None
    return ", ".join(vals)

def extract_top_titles(df: pd.DataFrame, metric_key: str, max_rank: int = 5) -> Dict[str, Any]:
    site_col = find_col(df, "site")
    metric_col = find_col(df, metric_key)

    if not site_col or not metric_col:
        return {}

    temp = df[[site_col, metric_col]].copy()
    temp[metric_col] = temp[metric_col].apply(safe_number)
    temp = temp.dropna(subset=[metric_col])

    temp = temp[
        temp[site_col].apply(
            lambda x: clean_text(x).lower() not in ["site", "publisher", "domain", "environment", "property", "placement"]
        )
    ]

    if temp.empty:
        return {}

    temp = temp.sort_values(by=metric_col, ascending=False).head(max_rank)

    prefix = {
        "ctr": "TOP_TITLES_CTR",
        "engagement_rate": "TOP_TITLES_ER",
        "vcr": "TOP_TITLES_VCR",
    }[metric_key]

    out = {}
    for idx, (_, row) in enumerate(temp.iterrows(), start=1):
        out[f"{prefix}_{idx}_NAME"] = clean_text(row[site_col])
        out[f"{prefix}_{idx}_VALUE"] = format_percent(row[metric_col])

    return out

def build_mapped_values(sheets: Dict[str, pd.DataFrame], filename: str = "") -> Dict[str, Any]:
    # Optional only - do not fail if missing
    rules = load_rules_master()
    rules_available = rules is not None

    detected = detect_tables(sheets)

    kpi_df = detected.get("campaign_kpi_summary")
    delivery_df = detected.get("campaign_delivery")
    site_df = detected.get("site")
    geo_df = detected.get("geo")
    format_df = detected.get("format")

    start_dt, end_dt = extract_dates(sheets, detected, filename)
    client_name = extract_client_name(sheets, detected, filename)

    mapped = {}

    if kpi_df is not None:
        mapped["CAMPAIGN_NAME"] = clean_text(extract_kpi_value(kpi_df, "campaign"))
        mapped["DELIVERED_IMPRESSIONS"] = format_number(safe_number(extract_kpi_value(kpi_df, "impressions")), 0)
        mapped["PERFORMANCE_CTR"] = format_percent(extract_kpi_value(kpi_df, "ctr"))
        mapped["PERFORMANCE_ENGAGEMENT_RATE"] = format_percent(extract_kpi_value(kpi_df, "engagement_rate"))
        mapped["PERFORMANCE_VCR"] = format_percent(extract_kpi_value(kpi_df, "vcr"))
        mapped["PERFORMANCE_ON_SCREEN"] = format_percent(extract_kpi_value(kpi_df, "on_screen"))

        spend_val = extract_kpi_value(kpi_df, "spend")
        mapped["CAMPAIGN_BUDGET"] = format_currency(safe_number(spend_val))
    else:
        mapped["CAMPAIGN_NAME"] = None
        mapped["DELIVERED_IMPRESSIONS"] = None
        mapped["PERFORMANCE_CTR"] = None
        mapped["PERFORMANCE_ENGAGEMENT_RATE"] = None
        mapped["PERFORMANCE_VCR"] = None
        mapped["PERFORMANCE_ON_SCREEN"] = None
        mapped["CAMPAIGN_BUDGET"] = None

    if delivery_df is not None:
        mapped["IO_OVERALL_IMPRESSIONS"] = format_number(safe_number(extract_kpi_value(delivery_df, "sold_paid_units")), 0)
        mapped["DELIVERED_OVERALL_AV_UNITS"] = format_number(safe_number(extract_kpi_value(delivery_df, "delivered_overall_av_units")), 0)
        mapped["DELIVERY_WITH_AV_PERCENT"] = format_percent(extract_kpi_value(delivery_df, "delivery_incl_av"))

        av_amount = safe_number(extract_kpi_value(delivery_df, "delivered_av_amount"))
        mapped["ADDED_VALUE_WORTH"] = format_currency(abs(av_amount) if av_amount is not None else None)

        mapped["ADDED_VALUE_IMPRESSIONS"] = mapped["DELIVERED_OVERALL_AV_UNITS"]
    else:
        mapped["IO_OVERALL_IMPRESSIONS"] = None
        mapped["DELIVERED_OVERALL_AV_UNITS"] = None
        mapped["DELIVERY_WITH_AV_PERCENT"] = None
        mapped["ADDED_VALUE_WORTH"] = None
        mapped["ADDED_VALUE_IMPRESSIONS"] = None

    mapped["CLIENT_NAME"] = client_name
    mapped["CAMPAIGN_FORMATS"] = join_dimension_values(format_df, "format") if format_df is not None else None
    mapped["CAMPAIGN_MARKETS"] = join_dimension_values(geo_df, "geo") if geo_df is not None else None
    mapped["LIVE_DATES_SHORT"] = format_date_short(start_dt, end_dt)
    mapped["LIVE_DATES_FULL"] = format_date_full(start_dt, end_dt)
    mapped["CAMPAIGN_PERIOD"] = format_quarter_from_date(end_dt)

    if site_df is not None:
        mapped.update(extract_top_titles(site_df, "ctr", 5))
        mapped.update(extract_top_titles(site_df, "engagement_rate", 5))
        mapped.update(extract_top_titles(site_df, "vcr", 5))

    for metric_prefix in ["TOP_TITLES_CTR", "TOP_TITLES_ER", "TOP_TITLES_VCR"]:
        for i in range(1, 6):
            mapped.setdefault(f"{metric_prefix}_{i}_NAME", None)
            mapped.setdefault(f"{metric_prefix}_{i}_VALUE", None)

    diagnostics = {
        "detected_tables": list(detected.keys()),
        "table_columns": {k: [str(c) for c in v.columns] for k, v in detected.items()},
        "rules_master_loaded": rules_available,
        "version": APP_VERSION,
    }

    return {
        "mapped_values": mapped,
        "diagnostics": diagnostics,
    }

# -------------------------
# PPT PLACEHOLDER ENGINE
# -------------------------
def replace_text(text: str, data: dict) -> str:
    if not text:
        return text
    for key, value in data.items():
        placeholder = f"{{{{{key}}}}}"
        text = text.replace(placeholder, "" if value is None else str(value))
    return text

def replace_in_shape(shape, data: dict):
    if hasattr(shape, "text_frame") and shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.text = replace_text(run.text, data)

    if hasattr(shape, "table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.text = replace_text(run.text, data)

def generate_ppt_from_template(template_path: Path, output_path: Path, data: dict):
    prs = Presentation(template_path)
    for slide in prs.slides:
        for shape in slide.shapes:
            replace_in_shape(shape, data)
    prs.save(output_path)

# -------------------------
# VALIDATE EOC
# -------------------------
@app.post("/validate-eoc")
async def validate_eoc(eoc_file: UploadFile = File(...)):
    try:
        path = BASE_DIR / f"eoc_{datetime.now().timestamp()}.xlsx"

        with open(path, "wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        sheets = pd.read_excel(path, sheet_name=None, header=None)
        result = build_mapped_values(sheets, filename=eoc_file.filename)

        return JSONResponse(content={
            "status": "validated",
            "version": APP_VERSION,
            **result
        })

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc(),
                "version": APP_VERSION
            }
        )

# -------------------------
# GENERATE EXEC SUMMARY
# -------------------------
@app.post("/generate-exec-summary")
async def generate_exec_summary(
    eoc_file: UploadFile = File(...),
    template_file: UploadFile = File(...)
):
    try:
        eoc_path = BASE_DIR / f"eoc_{datetime.now().timestamp()}.xlsx"
        template_path = BASE_DIR / f"template_{datetime.now().timestamp()}.pptx"

        with open(eoc_path, "wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        with open(template_path, "wb") as f:
            shutil.copyfileobj(template_file.file, f)

        sheets = pd.read_excel(eoc_path, sheet_name=None, header=None)
        result = build_mapped_values(sheets, filename=eoc_file.filename)
        mapped_data = result["mapped_values"]

        temp_output = OUTPUT_DIR / "temp_output.pptx"
        generate_ppt_from_template(
            template_path=template_path,
            output_path=temp_output,
            data=mapped_data
        )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        final = OUTPUT_DIR / f"Exec_Summary_Output_{timestamp}.pptx"
        shutil.copy2(temp_output, final)

        return FileResponse(
            path=str(final),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=final.name,
        )

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc(),
                "version": APP_VERSION
            }
        )
