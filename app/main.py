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

app = FastAPI(title="PCA Automation API", version="8.3.0")

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
    try:
        if pd.isna(value):
            return ""
    except:
        pass
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


# -------------------------
# NUMBER HELPERS
# -------------------------
def safe_number(value):
    try:
        if isinstance(value, str):
            value = value.replace(",", "").replace("£", "").replace("$", "").replace("€", "").replace("%", "")
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
# DATE FORMAT
# -------------------------
def format_date(date_str):
    try:
        match = re.search(r"\d{4}-\d{2}-\d{2}", str(date_str))
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
        row = df.iloc[i].tolist()

        score = 0
        for cell in row:
            cell_text = clean_text(cell).lower()
            for kw in keywords:
                if kw in cell_text:
                    score += 1

        if score >= 3:
            return i

    return 0


# -------------------------
# COLUMN ALIASES
# -------------------------
ALIASES = {
    "campaign": ["campaign", "campaign name"],
    "impressions": ["delivered impressions", "impressions"],
    "ctr": ["ctr", "click through rate"],
    "engagement": ["engagement rate", "er"],
    "vcr": ["vcr", "video completion rate"],
    "spend": ["spend", "actual spend"],
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
# CORE EXTRACTION
# -------------------------
def extract_data(sheets):
    best_df = None
    best_score = -1

    for name, df in sheets.items():
        if df is None or df.empty:
            continue

        raw_df = df.copy()
        header_row = find_header_row(raw_df)

        df = raw_df.iloc[header_row:].copy()
        df.columns = [clean_text(c) for c in df.iloc[0]]
        df = df[1:].reset_index(drop=True)

        cols = [normalize_header(c) for c in df.columns]

        score = 0
        if any("campaign" in c for c in cols): score += 2
        if any("impressions" in c for c in cols): score += 2
        if any("ctr" in c for c in cols): score += 2
        if any("engagement" in c for c in cols): score += 2
        if any("vcr" in c for c in cols): score += 2
        if any("spend" in c for c in cols): score += 2

        if score > best_score:
            best_score = score
            best_df = df

    if best_df is None:
        return {}

    df = best_df
    first_row = df.iloc[0]

    return {
        "CAMPAIGN_NAME": clean_text(first_row[find_col(df, ALIASES["campaign"])]) if find_col(df, ALIASES["campaign"]) else None,
        "DELIVERED_IMPRESSIONS": safe_number(first_row[find_col(df, ALIASES["impressions"])]) if find_col(df, ALIASES["impressions"]) else None,
        "CTR": safe_percent(first_row[find_col(df, ALIASES["ctr"])]) if find_col(df, ALIASES["ctr"]) else None,
        "ENGAGEMENT_RATE": safe_percent(first_row[find_col(df, ALIASES["engagement"])]) if find_col(df, ALIASES["engagement"]) else None,
        "VCR": safe_percent(first_row[find_col(df, ALIASES["vcr"])]) if find_col(df, ALIASES["vcr"]) else None,
        "SPEND": safe_number(first_row[find_col(df, ALIASES["spend"])]) if find_col(df, ALIASES["spend"]) else None,
    }


# -------------------------
# TOP TITLES
# -------------------------
def extract_top_titles(sheets):
    for name, df in sheets.items():
        if df is None or df.empty:
            continue

        raw_df = df.copy()
        header_row = find_header_row(raw_df)

        df = raw_df.iloc[header_row:].copy()
        df.columns = [clean_text(c) for c in df.iloc[0]]
        df = df[1:].reset_index(drop=True)

        site_col = find_col(df, ALIASES["site"])
        ctr_col = find_col(df, ALIASES["ctr"])
        er_col = find_col(df, ALIASES["engagement"])

        metric_col = ctr_col if ctr_col else er_col

        if not site_col or not metric_col:
            continue

        df = df[[site_col, metric_col]].copy()
        df[metric_col] = df[metric_col].apply(safe_number)
        df = df.dropna().sort_values(by=metric_col, ascending=False)

        results = {}

        for i in range(min(3, len(df))):
            row = df.iloc[i]
            results[f"TOP_TITLE_{i+1}_NAME"] = clean_text(row[site_col])
            results[f"TOP_TITLE_{i+1}_CTR"] = safe_percent(row[metric_col])

        return results

    return {}


# -------------------------
# DATE EXTRACTION
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
# PPT
# -------------------------
def replace_text(text, data):
    if not text:
        return text

    for k, v in data.items():
        text = text.replace(f"{{{{{k}}}}}", str(v) if v else "")

    return text


def process_slide(slide, data):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                for r in p.runs:
                    r.text = replace_text(r.text, data)


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

        sheets = pd.read_excel(path, sheet_name=None, header=None)

        data = extract_data(sheets)
        data["REPORT_DATE"] = extract_date(sheets)
        data.update(extract_top_titles(sheets))

        return JSONResponse(content={"status": "validated", "mapped_values": data})

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e), "trace": traceback.format_exc()})


# -------------------------
# GENERATE PPT
# -------------------------
@app.post("/generate-exec-summary")
async def generate_exec_summary(eoc_file: UploadFile = File(...), template_file: UploadFile = File(...)):
    try:
        eoc_path = BASE_DIR / f"eoc_{datetime.now().timestamp()}.xlsx"
        template_path = BASE_DIR / f"template_{datetime.now().timestamp()}.pptx"

        with open(eoc_path, "wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        with open(template_path, "wb") as f:
            shutil.copyfileobj(template_file.file, f)

        sheets = pd.read_excel(eoc_path, sheet_name=None, header=None)

        data = extract_data(sheets)
        data["REPORT_DATE"] = extract_date(sheets)
        data.update(extract_top_titles(sheets))

        output = OUTPUT_DIR / "output.pptx"
        generate_ppt(template_path, output, data)

        return FileResponse(path=str(output), media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename="Exec_Summary_Output.pptx")

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e), "trace": traceback.format_exc()})
