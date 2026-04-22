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

app = FastAPI(title="PCA Automation API", version="8.3.1")

BASE_DIR = Path("/tmp")
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

@app.get("/health")
def health():
    return {"status": "ok"}

# -------------------------
# HELPERS
# -------------------------
def clean_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except:
        pass
    return re.sub(r"\s+", " ", str(value)).strip()

def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()

def safe_number(value):
    try:
        if isinstance(value, str):
            value = value.replace(",", "").replace("£", "").replace("%", "")
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

def format_date(text):
    try:
        match = re.search(r"\d{4}-\d{2}-\d{2}", str(text))
        if match:
            dt = datetime.strptime(match.group(), "%Y-%m-%d")
            return dt.strftime("%d %b %Y")
    except:
        pass
    return None

# -------------------------
# HEADER DETECTION
# -------------------------
def find_header_row(df):
    keywords = ["campaign", "impressions", "ctr", "engagement", "vcr", "spend"]

    for i in range(min(20, len(df))):
        score = 0
        for cell in df.iloc[i]:
            text = clean_text(cell).lower()
            for kw in keywords:
                if kw in text:
                    score += 1
        if score >= 3:
            return i
    return 0

# -------------------------
# ALIASES
# -------------------------
ALIASES = {
    "campaign": ["campaign"],
    "impressions": ["impressions"],
    "ctr": ["ctr"],
    "engagement": ["engagement"],
    "vcr": ["vcr"],
    "spend": ["spend"],
    "site": ["site", "placement", "publisher"]
}

def find_col(df, aliases):
    for col in df.columns:
        norm = normalize_header(col)
        for alias in aliases:
            if normalize_header(alias) in norm:
                return col
    return None

# -------------------------
# CORE KPI EXTRACTION
# -------------------------
def extract_data(sheets):
    best_df = None
    best_score = -1

    for df in sheets.values():
        if df is None or df.empty:
            continue

        header = find_header_row(df)

        temp = df.iloc[header:].copy()
        temp.columns = [clean_text(c) for c in temp.iloc[0]]
        temp = temp[1:].reset_index(drop=True)

        cols = [normalize_header(c) for c in temp.columns]

        score = sum([
            any("campaign" in c for c in cols),
            any("impressions" in c for c in cols),
            any("ctr" in c for c in cols),
            any("engagement" in c for c in cols),
            any("vcr" in c for c in cols),
            any("spend" in c for c in cols),
        ])

        if score > best_score:
            best_score = score
            best_df = temp

    if best_df is None:
        return {}

    df = best_df
    row = df.iloc[0]

    return {
        "CAMPAIGN_NAME": clean_text(row[find_col(df, ALIASES["campaign"])]) if find_col(df, ALIASES["campaign"]) else None,
        "DELIVERED_IMPRESSIONS": safe_number(row[find_col(df, ALIASES["impressions"])]) if find_col(df, ALIASES["impressions"]) else None,
        "CTR": safe_percent(row[find_col(df, ALIASES["ctr"])]) if find_col(df, ALIASES["ctr"]) else None,
        "ENGAGEMENT_RATE": safe_percent(row[find_col(df, ALIASES["engagement"])]) if find_col(df, ALIASES["engagement"]) else None,
        "VCR": safe_percent(row[find_col(df, ALIASES["vcr"])]) if find_col(df, ALIASES["vcr"]) else None,
        "SPEND": safe_number(row[find_col(df, ALIASES["spend"])]) if find_col(df, ALIASES["spend"]) else None,
    }

# -------------------------
# 🔥 FIXED TOP TITLES
# -------------------------
def extract_top_titles(sheets):
    best_df = None
    best_score = -1

    for df in sheets.values():
        if df is None or df.empty:
            continue

        header = find_header_row(df)

        temp = df.iloc[header:].copy()
        temp.columns = [clean_text(c) for c in temp.iloc[0]]
        temp = temp[1:].reset_index(drop=True)

        cols = [normalize_header(c) for c in temp.columns]

        score = 0
        if any("site" in c for c in cols):
            score += 3
        if any("ctr" in c for c in cols):
            score += 2

        if score > best_score:
            best_score = score
            best_df = temp

    if best_df is None:
        return {}

    df = best_df

    site_col = find_col(df, ALIASES["site"])
    metric_col = find_col(df, ALIASES["ctr"]) or find_col(df, ALIASES["engagement"])

    if not site_col or not metric_col:
        return {}

    df = df[[site_col, metric_col]].copy()
    df[metric_col] = df[metric_col].apply(safe_number)

    df = df.dropna().sort_values(by=metric_col, ascending=False)

    result = {}
    for i in range(min(3, len(df))):
        row = df.iloc[i]
        result[f"TOP_TITLE_{i+1}_NAME"] = clean_text(row[site_col])
        result[f"TOP_TITLE_{i+1}_CTR"] = safe_percent(row[metric_col])

    return result

# -------------------------
# DATE
# -------------------------
def extract_date(sheets):
    for df in sheets.values():
        if df is None or df.empty:
            continue

        for row in df.head(15).values:
            for cell in row:
                text = clean_text(cell)
                if "report date" in text.lower():
                    return format_date(text)

    return None

# -------------------------
# VALIDATE
# -------------------------
@app.post("/validate-eoc")
async def validate_eoc(eoc_file: UploadFile = File(...)):
    try:
        path = BASE_DIR / f"eoc_{datetime.now().timestamp()}.xlsx"

        with open(path, "wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        sheets = pd.read_excel(path, sheet_name=None, header=None)

        data = extract_data(sheets)
        data["REPORT_DATE"] = extract_date(sheets)
        data.update(extract_top_titles(sheets))

        return JSONResponse(content={"status": "validated", "mapped_values": data})

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e), "trace": traceback.format_exc()})
