from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
import pandas as pd
import shutil
from pathlib import Path
from datetime import datetime
import traceback
import math
import re

APP_VERSION = "9.0.0"
app = FastAPI(title="PCA Automation API", version=APP_VERSION)

BASE_DIR = Path("/tmp")

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
def clean_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


def safe_number(value):
    try:
        if isinstance(value, str):
            value = (
                value.replace(",", "")
                .replace("£", "")
                .replace("$", "")
                .replace("€", "")
                .replace("%", "")
            )
        value = float(value)
        if math.isnan(value) or math.isinf(value):
            return None
        return value
    except Exception:
        return None


def safe_percent(value):
    try:
        if isinstance(value, str) and "%" in value:
            return value.strip()

        num = float(value)

        if num > 1:
            return f"{round(num, 2)}%"
        return f"{round(num * 100, 2)}%"
    except Exception:
        return None


def format_date(text):
    try:
        match = re.search(r"\d{4}-\d{2}-\d{2}", str(text))
        if match:
            dt = datetime.strptime(match.group(), "%Y-%m-%d")
            return dt.strftime("%d %b %Y")
    except Exception:
        pass
    return None


# -------------------------
# HEADER DETECTION
# -------------------------
def find_header_row(df):
    keywords = ["campaign", "impressions", "ctr", "engagement", "vcr", "completion", "spend"]

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
    "ctr": ["ctr", "click through rate"],
    "engagement": ["engagement rate", "er"],
    "vcr": ["vcr", "video completion rate", "completion rate"],
    "spend": ["spend", "media spend"],
}


def find_col(df, aliases):
    for col in df.columns:
        norm = normalize_header(col)
        for alias in aliases:
            if normalize_header(alias) in norm:
                return col
    return None


# -------------------------
# KPI EXTRACTION
# -------------------------
def extract_data(sheets):
    for df in sheets.values():
        if df is None or df.empty:
            continue

        header = find_header_row(df)

        temp = df.iloc[header:].copy()
        temp.columns = [clean_text(c) for c in temp.iloc[0]]
        temp = temp[1:].reset_index(drop=True)

        campaign_col = find_col(temp, ALIASES["campaign"])
        impressions_col = find_col(temp, ALIASES["impressions"])

        if campaign_col and impressions_col:
            row = temp.iloc[0]

            ctr_col = find_col(temp, ALIASES["ctr"])
            engagement_col = find_col(temp, ALIASES["engagement"])
            vcr_col = find_col(temp, ALIASES["vcr"])
            spend_col = find_col(temp, ALIASES["spend"])

            return {
                "CAMPAIGN_NAME": clean_text(row[campaign_col]),
                "DELIVERED_IMPRESSIONS": safe_number(row[impressions_col]),
                "CTR": safe_percent(row[ctr_col]) if ctr_col else None,
                "ENGAGEMENT_RATE": safe_percent(row[engagement_col]) if engagement_col else None,
                "VCR": safe_percent(row[vcr_col]) if vcr_col else None,
                "SPEND": safe_number(row[spend_col]) if spend_col else None,
            }

    return {}


# -------------------------
# TOP TITLES
# -------------------------
def extract_top_titles(sheets):
    for df in sheets.values():
        if df is None or df.empty:
            continue

        header = find_header_row(df)

        temp = df.iloc[header:].copy()
        temp.columns = [clean_text(c) for c in temp.iloc[0]]
        temp = temp[1:].reset_index(drop=True)

        site_col = None
        for col in temp.columns:
            if any(x in normalize_header(col) for x in ["site", "domain", "publisher", "environment", "property"]):
                site_col = col
                break

        if not site_col:
            continue

        metric_col = (
            find_col(temp, ALIASES["ctr"]) or
            find_col(temp, ALIASES["engagement"]) or
            find_col(temp, ALIASES["impressions"])
        )

        if not metric_col:
            continue

        temp = temp[[site_col, metric_col]].copy()
        temp[metric_col] = temp[metric_col].apply(safe_number)
        temp = temp.dropna().sort_values(by=metric_col, ascending=False)

        result = {}
        for i in range(min(3, len(temp))):
            row = temp.iloc[i]
            result[f"TOP_TITLE_{i+1}_NAME"] = clean_text(row[site_col])

            if "impression" in normalize_header(metric_col):
                result[f"TOP_TITLE_{i+1}_CTR"] = str(int(row[metric_col]))
            else:
                result[f"TOP_TITLE_{i+1}_CTR"] = safe_percent(row[metric_col])

        return result

    return {}


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

        return JSONResponse(content={
            "status": "validated",
            "mapped_values": data,
            "version": APP_VERSION
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
