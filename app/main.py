from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse, FileResponse
import pandas as pd
import shutil
from pathlib import Path
from datetime import datetime
import traceback
from pptx import Presentation
import math
import re

app = FastAPI(title="PCA Automation API", version="8.1.0")

# -------------------------
# CONFIG
# -------------------------
BASE_DIR = Path("/tmp")
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


# -------------------------
# HEALTH
# -------------------------
@app.get("/health")
def health():
    return {"status": "ok"}


# -------------------------
# SAFE HELPERS
# -------------------------
def clean_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_header(value):
    text = clean_text(value).lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def is_blank(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return clean_text(value) == ""


def safe_number_from_value(value):
    try:
        if value is None:
            return None
        if isinstance(value, str):
            value = value.replace(",", "").replace("%", "").replace("£", "").replace("$", "").replace("€", "").strip()
        value = float(value)
        if math.isnan(value) or math.isinf(value):
            return None
        return value
    except Exception:
        return None


def safe_json_value(value):
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()

    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None

    return value


def safe_percent_from_value(value):
    number = safe_number_from_value(value)
    if number is None:
        return None

    # If decimal like 0.0034 -> 0.34%
    # If already like 34.2 -> 34.2%
    if number <= 1:
        number = number * 100

    return f"{round(number, 2)}%"


def first_non_null(series):
    try:
        for val in series:
            if not is_blank(val):
                return val
    except Exception:
        pass
    return None


# -------------------------
# LIGHT HYBRID COLUMN MATCHING
# -------------------------
COLUMN_ALIASES = {
    "campaign_name": [
        "campaign", "campaign name", "name"
    ],
    "delivered_impressions": [
        "delivered impressions", "impressions", "served impressions"
    ],
    "ctr": [
        "ctr", "click through rate", "click-through rate"
    ],
    "engagement_rate": [
        "engagement rate", "total er", "er"
    ],
    "vcr": [
        "vcr", "video completion rate", "video completions rate"
    ],
    "spend": [
        "actual spend", "spend", "media spend", "total spend"
    ],
    "report_date": [
        "report date", "date", "live date", "campaign date"
    ],
}


def find_best_matching_column(columns, aliases):
    best_col = None
    best_score = 0

    normalized_columns = {col: normalize_header(col) for col in columns}

    for col, norm_col in normalized_columns.items():
        for alias in aliases:
            norm_alias = normalize_header(alias)

            if norm_col == norm_alias:
                score = 100
            elif norm_alias in norm_col:
                score = 80
            elif norm_col in norm_alias:
                score = 60
            else:
                score = 0

            if score > best_score:
                best_score = score
                best_col = col

    return best_col if best_score >= 60 else None


def get_value_from_df(df, aliases, as_percent=False, as_number=False):
    col = find_best_matching_column(df.columns, aliases)
    if not col:
        return None

    value = first_non_null(df[col])
    if value is None:
        return None

    if as_percent:
        return safe_percent_from_value(value)

    if as_number:
        return safe_number_from_value(value)

    return safe_json_value(value)


# -------------------------
# WORKBOOK / SHEET HELPERS
# -------------------------
def read_all_sheets(path):
    try:
        return pd.read_excel(path, sheet_name=None)
    except Exception:
        return {}


def choose_best_sheet(sheets):
    """
    Keep this simple:
    prefer sheet with most useful KPI-style columns.
    """
    best_name = None
    best_df = None
    best_score = -1

    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue

        score = 0
        cols = list(df.columns)

        if find_best_matching_column(cols, COLUMN_ALIASES["campaign_name"]):
            score += 2
        if find_best_matching_column(cols, COLUMN_ALIASES["delivered_impressions"]):
            score += 2
        if find_best_matching_column(cols, COLUMN_ALIASES["ctr"]):
            score += 2
        if find_best_matching_column(cols, COLUMN_ALIASES["engagement_rate"]):
            score += 2
        if find_best_matching_column(cols, COLUMN_ALIASES["vcr"]):
            score += 2
        if find_best_matching_column(cols, COLUMN_ALIASES["spend"]):
            score += 2

        if score > best_score:
            best_score = score
            best_name = sheet_name
            best_df = df

    return best_name, best_df


def scan_workbook_for_report_date(sheets):
    for _, df in sheets.items():
        if df is None or df.empty:
            continue

        # scan small area only
        preview = df.head(10)
        for _, row in preview.iterrows():
            for val in row.tolist():
                text = clean_text(val)
                if "report date" in text.lower():
                    return text
    return None


# -------------------------
# CORE EXTRACTION
# -------------------------
def extract_eoc_data_from_sheets(sheets: dict):
    result = {}

    selected_sheet_name, df = choose_best_sheet(sheets)

    if df is None or df.empty:
        return {
            "CAMPAIGN_NAME": None,
            "DELIVERED_IMPRESSIONS": None,
            "CTR": None,
            "ENGAGEMENT_RATE": None,
            "VCR": None,
            "SPEND": None,
            "REPORT_DATE": None,
            "_SELECTED_SHEET": None,
        }

    campaign_name = get_value_from_df(df, COLUMN_ALIASES["campaign_name"])
    delivered_impressions = get_value_from_df(df, COLUMN_ALIASES["delivered_impressions"], as_number=True)
    ctr = get_value_from_df(df, COLUMN_ALIASES["ctr"], as_percent=True)
    engagement_rate = get_value_from_df(df, COLUMN_ALIASES["engagement_rate"], as_percent=True)
    vcr = get_value_from_df(df, COLUMN_ALIASES["vcr"], as_percent=True)
    spend = get_value_from_df(df, COLUMN_ALIASES["spend"], as_number=True)
    report_date = get_value_from_df(df, COLUMN_ALIASES["report_date"])

    if report_date is None:
        report_date = scan_workbook_for_report_date(sheets)

    result["CAMPAIGN_NAME"] = campaign_name
    result["DELIVERED_IMPRESSIONS"] = delivered_impressions
    result["CTR"] = ctr
    result["ENGAGEMENT_RATE"] = engagement_rate
    result["VCR"] = vcr
    result["SPEND"] = spend
    result["REPORT_DATE"] = report_date
    result["_SELECTED_SHEET"] = selected_sheet_name

    return result


# -------------------------
# PLACEHOLDER ENGINE
# -------------------------
def replace_text(text, data):
    if not text:
        return text

    for key, value in data.items():
        if key.startswith("_"):
            continue
        placeholder = f"{{{{{key}}}}}"
        if placeholder in text:
            text = text.replace(placeholder, str(value) if value is not None else "")
    return text


def replace_in_shape(shape, data):
    if not shape.has_text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.text = replace_text(run.text, data)


def replace_in_table(table, data):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.text = replace_text(run.text, data)


def process_slide(slide, data):
    for shape in slide.shapes:
        if shape.has_text_frame:
            replace_in_shape(shape, data)

        if shape.has_table:
            replace_in_table(shape.table, data)


def generate_ppt(template_path: Path, output_path: Path, data: dict):
    prs = Presentation(template_path)

    for slide in prs.slides:
        process_slide(slide, data)

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

        sheets = read_all_sheets(path)
        mapped = extract_eoc_data_from_sheets(sheets)

        return JSONResponse(content={
            "status": "validated",
            "mapped_values": {
                k: safe_json_value(v) for k, v in mapped.items()
            }
        })

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc()
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
        # -------------------------
        # SAVE FILES
        # -------------------------
        eoc_path = BASE_DIR / f"eoc_{datetime.now().timestamp()}.xlsx"
        template_path = BASE_DIR / f"template_{datetime.now().timestamp()}.pptx"

        with open(eoc_path, "wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        with open(template_path, "wb") as f:
            shutil.copyfileobj(template_file.file, f)

        # -------------------------
        # EXTRACT DATA
        # -------------------------
        sheets = read_all_sheets(eoc_path)
        mapped_data = extract_eoc_data_from_sheets(sheets)

        # -------------------------
        # GENERATE PPT
        # -------------------------
        temp_output = OUTPUT_DIR / "temp_output.pptx"

        generate_ppt(
            template_path=template_path,
            output_path=temp_output,
            data=mapped_data
        )

        # -------------------------
        # FINAL OUTPUT
        # -------------------------
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
                "trace": traceback.format_exc()
            }
        )
