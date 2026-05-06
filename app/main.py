# main.py
# PCA Automation API
# v10.3.0 - Adds Slide Deck generation route while preserving stable One Pager flow

import os
import re
import io
import json
import math
import tempfile
import requests
from copy import deepcopy
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from pptx import Presentation


# ============================================================
# CONFIG
# ============================================================

APP_VERSION = "10.3.0-slide-deck"

MISSING_PPT_VALUE = "N/A"
MISSING_DISPLAY_VALUE = "N/A (not specified in source file)"
MAX_RETURN_FILE_BYTES = 10 * 1024 * 1024

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PATH = os.getenv(
    "TEMPLATE_PATH",
    os.path.join(BASE_DIR, "templates", "exec_summary_master.pptx")
)

SLIDE_DECK_TEMPLATE_PATH = os.getenv(
    "SLIDE_DECK_TEMPLATE_PATH",
    os.path.join(BASE_DIR, "templates", "exec_summary_slide_deck_master.pptx")
)

RULES_MASTER_PATH = os.getenv(
    "RULES_MASTER_PATH",
    os.path.join(BASE_DIR, "assets", "PCA_GPT_Rules_Master.xlsx")
)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")


# ============================================================
# APP
# ============================================================

app = FastAPI(
    title="PCA Automation API",
    version=APP_VERSION,
    description="Validates EOC Excel files and generates PCA PowerPoint outputs."
)


# ============================================================
# MODELS
# ============================================================

class OpenAIFileRef(BaseModel):
    id: Optional[str] = None
    file_id: Optional[str] = None
    name: Optional[str] = None
    filename: Optional[str] = None
    mime_type: Optional[str] = None
    download_link: Optional[str] = None
    url: Optional[str] = None


class ValidateEocRequest(BaseModel):
    openaiFileIdRefs: List[OpenAIFileRef] = Field(default_factory=list)


class GenerateExecSummaryRequest(BaseModel):
    openaiFileIdRefs: List[OpenAIFileRef] = Field(default_factory=list)
    exec_summary: Optional[str] = None


class GenerateSlideDeckRequest(BaseModel):
    openaiFileIdRefs: List[OpenAIFileRef] = Field(default_factory=list)
    exec_summary: Optional[str] = None


# ============================================================
# HELPERS
# ============================================================

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def is_blank(value: Any) -> bool:
    return clean_text(value) == ""


def normalise_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


def safe_display(value: Any) -> str:
    text = clean_text(value)
    return text if text else MISSING_DISPLAY_VALUE


def safe_ppt(value: Any) -> str:
    text = clean_text(value)
    return text if text else MISSING_PPT_VALUE


def format_int(value: Any) -> str:
    try:
        if value is None or value == "":
            return ""
        return f"{int(round(float(value))):,}"
    except Exception:
        return clean_text(value)


def format_currency(value: Any) -> str:
    try:
        if value is None or value == "":
            return ""
        return f"£{float(value):,.2f}".replace(".00", "")
    except Exception:
        return clean_text(value)


def format_percent(value: Any) -> str:
    try:
        if value is None or value == "":
            return ""
        number = float(value)
        if number <= 1:
            number = number * 100
        return f"{number:.2f}%"
    except Exception:
        return clean_text(value)


def slug_filename(value: str) -> str:
    text = clean_text(value)
    text = re.sub(r"[\\/:*?\"<>|]+", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or "Campaign"


def extract_file_id(ref: OpenAIFileRef) -> Optional[str]:
    return ref.file_id or ref.id


def download_openai_file(file_ref: OpenAIFileRef) -> bytes:
    if file_ref.download_link or file_ref.url:
        url = file_ref.download_link or file_ref.url
        response = requests.get(url, timeout=60)
        response.raise_for_status()
        return response.content

    file_id = extract_file_id(file_ref)

    if not file_id:
        raise HTTPException(status_code=400, detail="No OpenAI file id found in openaiFileIdRefs.")

    if not OPENAI_API_KEY:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY is not configured.")

    url = f"https://api.openai.com/v1/files/{file_id}/content"
    headers = {"Authorization": f"Bearer {OPENAI_API_KEY}"}

    response = requests.get(url, headers=headers, timeout=90)

    if response.status_code >= 400:
        raise HTTPException(
            status_code=500,
            detail=f"Could not download OpenAI file: {response.status_code}"
        )

    return response.content


def get_uploaded_excel_bytes(openai_file_refs: List[OpenAIFileRef]) -> bytes:
    if not openai_file_refs:
        raise HTTPException(status_code=400, detail="No file provided. Please upload an EOC Excel file.")

    return download_openai_file(openai_file_refs[0])


# ============================================================
# EOC PARSING
# ============================================================

def read_excel_sheets(excel_bytes: bytes) -> Dict[str, pd.DataFrame]:
    try:
        with io.BytesIO(excel_bytes) as buffer:
            sheets = pd.read_excel(buffer, sheet_name=None, header=None)
        return sheets
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read Excel file: {str(e)}")


def dataframe_to_text_grid(sheets: Dict[str, pd.DataFrame]) -> List[List[str]]:
    rows = []
    for sheet_name, df in sheets.items():
        rows.append([f"SHEET: {sheet_name}"])
        for _, row in df.iterrows():
            rows.append([clean_text(cell) for cell in row.tolist()])
    return rows


def find_value_near_label(rows: List[List[str]], aliases: List[str]) -> str:
    alias_norms = [normalise_key(a) for a in aliases]

    for row in rows:
        norm_cells = [normalise_key(c) for c in row]

        for idx, cell_norm in enumerate(norm_cells):
            if any(alias in cell_norm or cell_norm in alias for alias in alias_norms if alias):
                for next_idx in range(idx + 1, len(row)):
                    value = clean_text(row[next_idx])
                    if value and normalise_key(value) not in alias_norms:
                        return value

    return ""


def find_dates(rows: List[List[str]]) -> Dict[str, str]:
    all_text = " ".join(" ".join(row) for row in rows)

    date_patterns = [
        r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
        r"(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})",
        r"([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})"
    ]

    found = []
    for pattern in date_patterns:
        found.extend(re.findall(pattern, all_text))

    if len(found) >= 2:
        return {
            "LIVE_DATES_SHORT": f"{found[0]} - {found[1]}",
            "LIVE_DATES_FULL": f"{found[0]} - {found[1]}",
            "CAMPAIGN_PERIOD": derive_quarter_from_dates(found[0], found[1])
        }

    return {
        "LIVE_DATES_SHORT": "",
        "LIVE_DATES_FULL": "",
        "CAMPAIGN_PERIOD": ""
    }


def parse_any_date(value: str) -> Optional[datetime]:
    value = clean_text(value)

    formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
        "%d %B %Y", "%d %b %Y",
        "%B %d %Y", "%b %d %Y",
        "%B %d, %Y", "%b %d, %Y"
    ]

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except Exception:
            pass

    return None


def derive_quarter_from_dates(start_value: str, end_value: str) -> str:
    end_date = parse_any_date(end_value) or parse_any_date(start_value)

    if not end_date:
        return ""

    quarter = ((end_date.month - 1) // 3) + 1
    return f"Q{quarter} {end_date.year}"


def find_best_header_table(rows: List[List[str]]) -> pd.DataFrame:
    header_keywords = [
        "site", "publisher", "title", "impressions", "ctr", "engagement",
        "video completion", "vcr", "on screen", "viewability", "spend",
        "budget", "delivered"
    ]

    best_idx = None
    best_score = 0

    for i, row in enumerate(rows):
        norm = " ".join(normalise_key(c) for c in row)
        score = sum(1 for keyword in header_keywords if keyword in norm)

        if score > best_score:
            best_score = score
            best_idx = i

    if best_idx is None or best_score < 2:
        return pd.DataFrame()

    header = [clean_text(c) or f"Column_{i}" for i, c in enumerate(rows[best_idx])]
    data = []

    for row in rows[best_idx + 1:]:
        if all(is_blank(c) for c in row):
            if data:
                break
            continue

        if len(row) < len(header):
            row = row + [""] * (len(header) - len(row))

        data.append(row[:len(header)])

    if not data:
        return pd.DataFrame()

    return pd.DataFrame(data, columns=header)


def find_column(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    alias_norms = [normalise_key(a) for a in aliases]

    for col in df.columns:
        col_norm = normalise_key(col)
        if any(alias in col_norm or col_norm in alias for alias in alias_norms):
            return col

    return None


def numeric_series(df: pd.DataFrame, col: Optional[str]) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series(dtype=float)

    return pd.to_numeric(df[col], errors="coerce")


def top_performers(df: pd.DataFrame, metric_aliases: List[str], name_aliases: List[str]) -> List[Dict[str, str]]:
    if df.empty:
        return []

    name_col = find_column(df, name_aliases)
    metric_col = find_column(df, metric_aliases)

    if not name_col or not metric_col:
        return []

    temp = df[[name_col, metric_col]].copy()
    temp["_metric"] = pd.to_numeric(temp[metric_col], errors="coerce")
    temp = temp.dropna(subset=["_metric"])

    temp = temp[
        temp[name_col].astype(str).str.strip().ne("")
        & ~temp[name_col].astype(str).str.lower().str.contains("total")
    ]

    temp = temp.sort_values("_metric", ascending=False).head(3)

    results = []
    for _, row in temp.iterrows():
        results.append({
            "name": clean_text(row[name_col]),
            "value": format_percent(row["_metric"])
        })

    return results


def parse_eoc(excel_bytes: bytes) -> Dict[str, Any]:
    sheets = read_excel_sheets(excel_bytes)
    rows = dataframe_to_text_grid(sheets)
    df = find_best_header_table(rows)

    campaign_name = find_value_near_label(rows, ["Campaign Name", "Campaign", "Campaign Title"])
    client = find_value_near_label(rows, ["Client", "Advertiser", "Brand"])
    market = find_value_near_label(rows, ["Market", "Markets", "Geo", "Country"])
    creative_format = find_value_near_label(rows, ["Creative Format", "Format", "Formats"])

    date_values = find_dates(rows)

    site_aliases = ["Site", "Publisher", "Title", "Domain"]

    impressions_col = find_column(df, ["Delivered Impressions", "Impressions", "Delivered"])
    io_col = find_column(df, ["IO Overall Impressions", "Booked Impressions", "Sold Paid Units", "IO Impressions"])
    av_col = find_column(df, ["Delivered Overall AV", "Added Value Impressions", "Added Value Imps", "AV Units"])
    delivery_col = find_column(df, ["Delivery Percentage incl AV", "Delivery incl AV", "Delivery"])
    av_worth_col = find_column(df, ["Delivered AV Amount", "Added Value Worth", "AV Worth"])
    budget_col = find_column(df, ["Spend", "Budget", "Campaign Budget"])

    ctr_col = find_column(df, ["CTR", "Click Through Rate"])
    engagement_col = find_column(df, ["Engagement Rate", "ER"])
    vcr_col = find_column(df, ["Video Completion Rate", "VCR"])
    onscreen_col = find_column(df, ["Mobkoi On Screen", "On Screen", "MRC Viewability", "Viewability"])

    delivered_impressions = format_int(numeric_series(df, impressions_col).sum()) if impressions_col else ""
    io_impressions = format_int(numeric_series(df, io_col).sum()) if io_col else ""
    av_impressions = format_int(numeric_series(df, av_col).sum()) if av_col else ""
    delivery_incl_av = format_percent(numeric_series(df, delivery_col).mean()) if delivery_col else ""
    av_worth = format_currency(numeric_series(df, av_worth_col).sum()) if av_worth_col else ""
    budget = format_currency(numeric_series(df, budget_col).sum()) if budget_col else ""

    ctr = format_percent(numeric_series(df, ctr_col).mean()) if ctr_col else ""
    engagement_rate = format_percent(numeric_series(df, engagement_col).mean()) if engagement_col else ""
    vcr = format_percent(numeric_series(df, vcr_col).mean()) if vcr_col else ""
    onscreen = format_percent(numeric_series(df, onscreen_col).mean()) if onscreen_col else ""

    top_ctr = top_performers(df, ["CTR", "Click Through Rate"], site_aliases)
    top_engagement = top_performers(df, ["Engagement Rate", "ER"], site_aliases)
    top_vcr = top_performers(df, ["Video Completion Rate", "VCR"], site_aliases)

    mapped_values = {
        "CAMPAIGN_NAME": safe_ppt(campaign_name),
        "CLIENT": safe_ppt(client),
        "MARKET": safe_ppt(market),
        "MARKETS": safe_ppt(market),
        "LIVE_DATES_SHORT": safe_ppt(date_values.get("LIVE_DATES_SHORT")),
        "LIVE_DATES_FULL": safe_ppt(date_values.get("LIVE_DATES_FULL")),
        "CAMPAIGN_PERIOD": safe_ppt(date_values.get("CAMPAIGN_PERIOD")),

        "DELIVERED_IMPRESSIONS": safe_ppt(delivered_impressions),
        "IO_OVERALL_IMPRESSIONS": safe_ppt(io_impressions),
        "ADDED_VALUE_IMPRESSIONS": safe_ppt(av_impressions),
        "DELIVERY_INCL_AV": safe_ppt(delivery_incl_av),
        "ADDED_VALUE_WORTH": safe_ppt(av_worth),
        "CAMPAIGN_BUDGET": safe_ppt(budget),
        "BUDGET": safe_ppt(budget),

        "PERFORMANCE_CTR": safe_ppt(ctr),
        "PERFORMANCE_ENGAGEMENT_RATE": safe_ppt(engagement_rate),
        "PERFORMANCE_VCR": safe_ppt(vcr),
        "PERFORMANCE_VIEWABILITY": safe_ppt(onscreen),
        "PERFORMANCE_ON_SCREEN": safe_ppt(onscreen),

        "CREATIVE_FORMAT": safe_ppt(creative_format),
        "FORMAT": safe_ppt(creative_format),
    }

    for i in range(3):
        rank = i + 1

        mapped_values[f"TOP_TITLES_CTR_{rank}_NAME"] = safe_ppt(top_ctr[i]["name"] if i < len(top_ctr) else "")
        mapped_values[f"TOP_TITLES_CTR_{rank}_VALUE"] = safe_ppt(top_ctr[i]["value"] if i < len(top_ctr) else "")

        mapped_values[f"TOP_TITLES_ENGAGEMENT_{rank}_NAME"] = safe_ppt(top_engagement[i]["name"] if i < len(top_engagement) else "")
        mapped_values[f"TOP_TITLES_ENGAGEMENT_{rank}_VALUE"] = safe_ppt(top_engagement[i]["value"] if i < len(top_engagement) else "")

        mapped_values[f"TOP_TITLES_VCR_{rank}_NAME"] = safe_ppt(top_vcr[i]["name"] if i < len(top_vcr) else "")
        mapped_values[f"TOP_TITLES_VCR_{rank}_VALUE"] = safe_ppt(top_vcr[i]["value"] if i < len(top_vcr) else "")

    display_values = {
        key: value if value != MISSING_PPT_VALUE else MISSING_DISPLAY_VALUE
        for key, value in mapped_values.items()
    }

    return {
        "display_values": display_values,
        "mapped_values": mapped_values,
        "top_performers": {
            "ctr": top_ctr,
            "engagement_rate": top_engagement,
            "vcr": top_vcr
        },
        "detected_columns": {
            "impressions": impressions_col,
            "io_impressions": io_col,
            "added_value": av_col,
            "delivery_incl_av": delivery_col,
            "av_worth": av_worth_col,
            "budget": budget_col,
            "ctr": ctr_col,
            "engagement_rate": engagement_col,
            "vcr": vcr_col,
            "on_screen": onscreen_col
        }
    }


# ============================================================
# POWERPOINT PLACEHOLDER REPLACEMENT
# ============================================================

def replace_text_preserve_runs(paragraph, replacements: Dict[str, str]) -> None:
    full_text = "".join(run.text for run in paragraph.runs)

    if "{{" not in full_text:
        return

    new_text = full_text

    for key, value in replacements.items():
        new_text = new_text.replace(f"{{{{{key}}}}}", safe_ppt(value))

    unresolved = re.findall(r"\{\{[^}]+\}\}", new_text)
    for placeholder in unresolved:
        new_text = new_text.replace(placeholder, MISSING_PPT_VALUE)

    if new_text == full_text:
        return

    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""


def replace_shape_text(shape, replacements: Dict[str, str]) -> None:
    if hasattr(shape, "text_frame") and shape.text_frame:
        for paragraph in shape.text_frame.paragraphs:
            replace_text_preserve_runs(paragraph, replacements)

    if hasattr(shape, "table"):
        for row in shape.table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    replace_text_preserve_runs(paragraph, replacements)

    if hasattr(shape, "shapes"):
        for sub_shape in shape.shapes:
            replace_shape_text(sub_shape, replacements)


def fill_ppt_template(template_path: str, mapped_values: Dict[str, str]) -> bytes:
    if not os.path.exists(template_path):
        raise HTTPException(status_code=500, detail=f"Template not found: {template_path}")

    prs = Presentation(template_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_shape_text(shape, mapped_values)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    ppt_bytes = output.read()

    if len(ppt_bytes) > MAX_RETURN_FILE_BYTES:
        raise HTTPException(status_code=500, detail="Generated PPT exceeds maximum return size.")

    return ppt_bytes


def encode_file_response(filename: str, content: bytes) -> Dict[str, Any]:
    import base64

    return {
        "filename": filename,
        "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "content_base64": base64.b64encode(content).decode("utf-8")
    }


# ============================================================
# API RESPONSE BUILDERS
# ============================================================

def build_validation_summary(parsed: Dict[str, Any]) -> Dict[str, Any]:
    display = parsed["display_values"]
    top = parsed["top_performers"]

    return {
        "status": "success",
        "version": APP_VERSION,
        "display_values": display,
        "top_performers": top,
        "validation_summary": {
            "campaign_name": display.get("CAMPAIGN_NAME", MISSING_DISPLAY_VALUE),
            "client": display.get("CLIENT", MISSING_DISPLAY_VALUE),
            "market": display.get("MARKET", MISSING_DISPLAY_VALUE),
            "live_dates": display.get("LIVE_DATES_SHORT", MISSING_DISPLAY_VALUE),
            "campaign_period": display.get("CAMPAIGN_PERIOD", MISSING_DISPLAY_VALUE),
            "delivered_impressions": display.get("DELIVERED_IMPRESSIONS", MISSING_DISPLAY_VALUE),
            "io_overall_impressions": display.get("IO_OVERALL_IMPRESSIONS", MISSING_DISPLAY_VALUE),
            "added_value_impressions": display.get("ADDED_VALUE_IMPRESSIONS", MISSING_DISPLAY_VALUE),
            "delivery_incl_av": display.get("DELIVERY_INCL_AV", MISSING_DISPLAY_VALUE),
            "added_value_worth": display.get("ADDED_VALUE_WORTH", MISSING_DISPLAY_VALUE),
            "budget": display.get("CAMPAIGN_BUDGET", MISSING_DISPLAY_VALUE),
            "ctr": display.get("PERFORMANCE_CTR", MISSING_DISPLAY_VALUE),
            "engagement_rate": display.get("PERFORMANCE_ENGAGEMENT_RATE", MISSING_DISPLAY_VALUE),
            "vcr": display.get("PERFORMANCE_VCR", MISSING_DISPLAY_VALUE),
            "on_screen_rate": display.get("PERFORMANCE_ON_SCREEN", MISSING_DISPLAY_VALUE),
            "creative_format": display.get("CREATIVE_FORMAT", MISSING_DISPLAY_VALUE)
        },
        "detected_columns": parsed.get("detected_columns", {})
    }


def build_output_filename(mapped_values: Dict[str, str], output_type: str) -> str:
    campaign_name = mapped_values.get("CAMPAIGN_NAME", "Campaign")
    campaign_name = campaign_name if campaign_name != MISSING_PPT_VALUE else "Campaign"

    clean_campaign = slug_filename(campaign_name)

    if output_type == "slides":
        return f"Exec Summary_PCA Slides_{clean_campaign}.pptx"

    return f"Exec Summary_PCA One Pager_{clean_campaign}.pptx"


# ============================================================
# ROUTES
# ============================================================

@app.get("/health")
def health():
    return {
        "status": "ok",
        "version": APP_VERSION,
        "one_pager_template": TEMPLATE_PATH,
        "slide_deck_template": SLIDE_DECK_TEMPLATE_PATH,
        "rules_master": RULES_MASTER_PATH
    }


@app.post("/validate-eoc")
def validate_eoc(request: ValidateEocRequest):
    excel_bytes = get_uploaded_excel_bytes(request.openaiFileIdRefs)
    parsed = parse_eoc(excel_bytes)
    return JSONResponse(build_validation_summary(parsed))


@app.post("/generate-exec-summary")
def generate_exec_summary(request: GenerateExecSummaryRequest):
    excel_bytes = get_uploaded_excel_bytes(request.openaiFileIdRefs)
    parsed = parse_eoc(excel_bytes)

    mapped_values = parsed["mapped_values"]
    mapped_values["EXEC_SUMMARY"] = safe_ppt(request.exec_summary)

    ppt_bytes = fill_ppt_template(TEMPLATE_PATH, mapped_values)
    filename = build_output_filename(mapped_values, "one_pager")

    return JSONResponse({
        "status": "success",
        "version": APP_VERSION,
        "output_type": "one_pager",
        "filename": filename,
        "file": encode_file_response(filename, ppt_bytes)
    })


@app.post("/generate-slide-deck")
def generate_slide_deck(request: GenerateSlideDeckRequest):
    excel_bytes = get_uploaded_excel_bytes(request.openaiFileIdRefs)
    parsed = parse_eoc(excel_bytes)

    mapped_values = parsed["mapped_values"]
    mapped_values["EXEC_SUMMARY"] = safe_ppt(request.exec_summary)

    ppt_bytes = fill_ppt_template(SLIDE_DECK_TEMPLATE_PATH, mapped_values)
    filename = build_output_filename(mapped_values, "slides")

    return JSONResponse({
        "status": "success",
        "version": APP_VERSION,
        "output_type": "slides",
        "filename": filename,
        "file": encode_file_response(filename, ppt_bytes)
    })


# ============================================================
# LOCAL DEV
# ============================================================

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=int(os.getenv("PORT", 8000)),
        reload=True
    )
