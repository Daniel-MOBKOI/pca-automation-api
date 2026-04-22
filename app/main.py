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

app = FastAPI(title="PCA Automation API", version="8.2.0")

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
# TEXT HELPERS
# -------------------------
def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


# -------------------------
# NUMBER / FORMAT HELPERS
# -------------------------
def safe_number(value):
    try:
        if isinstance(value, str):
            value = value.replace(",", "").replace("£", "").replace("$", "").replace("%", "")
        value = float(value)
        if math.isnan(value) or math.isinf(value):
            return None
        return value
    except:
        return None


def safe_percent(value):
    num = safe_number(value)
    if num is None:
        return None
    if num <= 1:
        num *= 100
    return f"{round(num, 2)}%"


# -------------------------
# DATE FORMAT (RULE FIX)
# -------------------------
def format_date(date_str):
    try:
        # Extract date like 2026-03-19
        match = re.search(r"\d{4}-\d{2}-\d{2}", str(date_str))
        if match:
            dt = datetime.strptime(match.group(), "%Y-%m-%d")
            return dt.strftime("%d %b %Y")
    except:
        pass
    return None


# -------------------------
# HEADER DETECTION (FIX)
# -------------------------
def find_header_row(df):
    keywords = ["campaign", "impressions", "ctr", "engagement", "vcr", "spend"]

    for i in range(min(15, len(df))):
        row = df.iloc[i].astype(str).str.lower()

        score = 0
        for cell in row:
            for kw in keywords:
                if kw in cell:
                    score += 1

        if score >= 3:
            return i

    return 0


# -------------------------
# COLUMN MATCHING
# -------------------------
ALIASES = {
    "campaign": ["campaign", "campaign name"],
    "impressions": ["delivered impressions", "impressions"],
    "ctr": ["ctr"],
    "engagement": ["engagement rate", "er"],
    "vcr": ["vcr", "video completion rate"],
    "spend": ["spend", "actual spend"],
    "date": ["report date", "date"]
}


def find_col(df, aliases):
    for col in df.columns:
        norm = normalize_header(col)
        for alias in aliases:
            if alias in norm:
                return col
    return None


# -------------------------
# CORE EXTRACTION
# -------------------------
def extract_data(sheets):
    best_df = None
    best_score = -1

    for name, df in sheets.items():
        if df is None or df.empty:
            continue

        header_row = find_header_row(df)

        df = df.iloc[header_row:]
        df.columns = df.iloc[0]
        df = df[1:]

        score = 0
        cols = [normalize_header(c) for c in df.columns]

        if any("campaign" in c for c in cols): score += 2
        if any("impressions" in c for c in cols): score += 2
        if any("ctr" in c for c in cols): score += 2

        if score > best_score:
            best_score = score
            best_df = df

    if best_df is None:
        return {}

    df = best_df

    campaign_col = find_col(df, ALIASES["campaign"])
    impressions_col = find_col(df, ALIASES["impressions"])
    ctr_col = find_col(df, ALIASES["ctr"])
    engagement_col = find_col(df, ALIASES["engagement"])
    vcr_col = find_col(df, ALIASES["vcr"])
    spend_col = find_col(df, ALIASES["spend"])

    first_row = df.iloc[0]

    return {
        "CAMPAIGN_NAME": clean_text(first_row[campaign_col]) if campaign_col else None,
        "DELIVERED_IMPRESSIONS": safe_number(first_row[impressions_col]) if impressions_col else None,
        "CTR": safe_percent(first_row[ctr_col]) if ctr_col else None,
        "ENGAGEMENT_RATE": safe_percent(first_row[engagement_col]) if engagement_col else None,
        "VCR": safe_percent(first_row[vcr_col]) if vcr_col else None,
        "SPEND": safe_number(first_row[spend_col]) if spend_col else None,
    }


# -------------------------
# DATE EXTRACTION (FIXED)
# -------------------------
def extract_date(sheets):
    for df in sheets.values():
        if df is None:
            continue

        for row in df.head(10).values:
            for cell in row:
                text = clean_text(cell)
                if "report date" in text.lower():
                    return format_date(text)

    return None


# -------------------------
# PPT FUNCTIONS (UNCHANGED)
# -------------------------
def replace_text(text, data):
    if not text:
        return text

    for key, value in data.items():
        placeholder = f"{{{{{key}}}}}"
        text = text.replace(placeholder, str(value) if value else "")

    return text


def process_slide(slide, data):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                for run in p.runs:
                    run.text = replace_text(run.text, data)


def generate_ppt(template_path, output_path, data):
    prs = Presentation(template_path)

    for slide in prs.slides:
        process_slide(slide, data)

    prs.save(output_path)


# -------------------------
# VALIDATE
# -------------------------
@app.post("/validate-eoc")
async def validate_eoc(eoc_file: UploadFile = File(...)):
    try:
        path = BASE_DIR / f"eoc_{datetime.now().timestamp()}.xlsx"

        with open(path, "wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        sheets = pd.read_excel(path, sheet_name=None)

        data = extract_data(sheets)
        data["REPORT_DATE"] = extract_date(sheets)

        return JSONResponse(content={
            "status": "validated",
            "mapped_values": data
        })

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e), "trace": traceback.format_exc()}
        )


# -------------------------
# GENERATE PPT
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

        sheets = pd.read_excel(eoc_path, sheet_name=None)

        data = extract_data(sheets)
        data["REPORT_DATE"] = extract_date(sheets)

        output = OUTPUT_DIR / "output.pptx"
        generate_ppt(template_path, output, data)

        return FileResponse(output)

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e), "trace": traceback.format_exc()}
        )
