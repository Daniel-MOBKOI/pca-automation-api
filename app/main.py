"""
PCA Automation Generator - Modular Builder Stable
Version: v12.1.0-modular-deck-filtering

Adds:
- Modular Slide Deck filtering
- SECTION_ID slide filtering
- SECTION_ID text box removal
- Deck modes: full / matching / custom
- Reuse One Pager section selections
- Preserve current stable One Pager workflow

Expected Render/GitHub structure:
- /templates/exec_summary_master.pptx
- /templates/exec_summary_slide_deck_master.pptx
- /assets/PCA_GPT_Rules_Master.xlsx

Core API endpoints:
- GET  /health
- POST /validate-eoc
- POST /generate-exec-summary
- POST /generate-modular-one-pager
- POST /generate-slide-deck
- POST /generate-one-pager-and-slides
"""

from __future__ import annotations

import base64
import copy
import io
import json
import os
import re
import tempfile
import uuid
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from pydantic import BaseModel, Field
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from starlette.responses import JSONResponse


# =============================================================================
# Config
# =============================================================================

APP_VERSION = "v12.1.0-modular-deck-filtering"

ROOT_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = Path(os.getenv("TEMPLATE_PATH", ROOT_DIR / "templates" / "exec_summary_master.pptx"))
SLIDE_DECK_TEMPLATE_PATH = Path(
    os.getenv("SLIDE_DECK_TEMPLATE_PATH", ROOT_DIR / "templates" / "exec_summary_slide_deck_master.pptx")
)
RULES_MASTER_PATH = Path(os.getenv("RULES_MASTER_PATH", ROOT_DIR / "assets" / "PCA_GPT_Rules_Master.xlsx"))

MAX_RETURN_FILE_BYTES = int(os.getenv("MAX_RETURN_FILE_BYTES", str(10 * 1024 * 1024)))
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

MISSING_PPT_VALUE = "N/A"
MISSING_DISPLAY_VALUE = "N/A (not specified in source file)"

SECTION_ID_PREFIX = "SECTION_ID:"

# Canonical modular section IDs.
SECTION_ALIASES = {
    "title overview": "TITLE_OVERVIEW",
    "title_overview": "TITLE_OVERVIEW",
    "exec summary": "TITLE_OVERVIEW",
    "executive summary": "TITLE_OVERVIEW",
    "overview": "TITLE_OVERVIEW",
    "title performance": "TITLE_PERFORMANCE",
    "title_performance": "TITLE_PERFORMANCE",
    "market performance": "MARKET_PERFORMANCE",
    "market_performance": "MARKET_PERFORMANCE",
    "markets": "MARKET_PERFORMANCE",
    "top titles markets": "TOP_TITLES_MARKETS",
    "top_titles_markets": "TOP_TITLES_MARKETS",
    "titles markets": "TOP_TITLES_MARKETS",
    "creative overview": "CREATIVE_OVERVIEW",
    "creative_overview": "CREATIVE_OVERVIEW",
    "creative performance": "CREATIVE_PERFORMANCE",
    "creative_performance": "CREATIVE_PERFORMANCE",
    "attention score": "ATTENTION_SCORE",
    "attention_score": "ATTENTION_SCORE",
    "attention": "ATTENTION_SCORE",
    "happydemics": "HAPPYDEMICS",
    "brand study": "BRAND_STUDY",
    "brand_study": "BRAND_STUDY",
    "lumen results": "LUMEN_RESULTS",
    "lumen_results": "LUMEN_RESULTS",
    "lumen learnings": "LUMEN_LEARNINGS",
    "lumen_learnings": "LUMEN_LEARNINGS",
    "learning recommendations": "LEARNING_RECOMMENDATIONS",
    "learning_recommendations": "LEARNING_RECOMMENDATIONS",
    "learnings": "LEARNING_RECOMMENDATIONS",
}

DEFAULT_BASE_SECTIONS = ["TITLE_OVERVIEW"]


# =============================================================================
# FastAPI
# =============================================================================

app = FastAPI(
    title="PCA Automation Generator",
    version=APP_VERSION,
    description="Validate EOC Excel files and generate PCA One Pagers / filtered Slide Decks.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =============================================================================
# Request / Response Models
# =============================================================================

class OpenAIFileRef(BaseModel):
    id: Optional[str] = None
    name: Optional[str] = None
    mime_type: Optional[str] = None
    download_link: Optional[str] = None
    data: Optional[str] = None


class ValidateEocRequest(BaseModel):
    openaiFileIdRefs: Optional[List[OpenAIFileRef]] = None
    file_base64: Optional[str] = None
    filename: Optional[str] = None


class GenerateExecSummaryRequest(BaseModel):
    openaiFileIdRefs: Optional[List[OpenAIFileRef]] = None
    file_base64: Optional[str] = None
    filename: Optional[str] = None
    exec_summary: Optional[str] = None
    section_selections: Optional[List[str]] = None
    one_pager_sections: Optional[List[str]] = None
    output_filename: Optional[str] = None


class GenerateSlideDeckRequest(BaseModel):
    openaiFileIdRefs: Optional[List[OpenAIFileRef]] = None
    file_base64: Optional[str] = None
    filename: Optional[str] = None
    exec_summary: Optional[str] = None
    section_selections: Optional[List[str]] = None
    one_pager_sections: Optional[List[str]] = None
    deck_mode: str = Field(default="matching", description="full, matching, or custom")
    custom_deck_sections: Optional[List[str]] = None
    output_filename: Optional[str] = None


class GenerateOnePagerAndSlidesRequest(BaseModel):
    openaiFileIdRefs: Optional[List[OpenAIFileRef]] = None
    file_base64: Optional[str] = None
    filename: Optional[str] = None
    exec_summary: Optional[str] = None
    section_selections: Optional[List[str]] = None
    one_pager_sections: Optional[List[str]] = None
    deck_mode: str = "matching"
    custom_deck_sections: Optional[List[str]] = None
    one_pager_filename: Optional[str] = None
    slide_deck_filename: Optional[str] = None


# =============================================================================
# Utility helpers
# =============================================================================

def now_stamp() -> str:
    return datetime.now().strftime("%d%m%Y")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def safe_display(value: Any) -> str:
    text = clean_text(value)
    return text if text else MISSING_DISPLAY_VALUE


def safe_ppt(value: Any) -> str:
    text = clean_text(value)
    return text if text else MISSING_PPT_VALUE


def slug_filename(text: str, fallback: str = "Campaign") -> str:
    text = clean_text(text) or fallback
    text = re.sub(r"[^A-Za-z0-9 _.-]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:80] or fallback


def canonical_section_id(section: str) -> str:
    raw = clean_text(section)
    if not raw:
        return ""
    raw = raw.replace(SECTION_ID_PREFIX, "")
    raw = raw.strip()
    upper = re.sub(r"[^A-Za-z0-9]+", "_", raw).strip("_").upper()
    alias_key = raw.lower().strip()
    return SECTION_ALIASES.get(alias_key, upper)


def canonical_section_list(sections: Optional[Sequence[str]], include_base: bool = True) -> List[str]:
    seen = set()
    output: List[str] = []

    if include_base:
        for item in DEFAULT_BASE_SECTIONS:
            if item not in seen:
                seen.add(item)
                output.append(item)

    for section in sections or []:
        sid = canonical_section_id(section)
        if sid and sid not in seen:
            seen.add(sid)
            output.append(sid)
    return output


def normalise_deck_mode(deck_mode: Optional[str]) -> str:
    mode = clean_text(deck_mode).lower()
    if mode not in {"full", "matching", "custom"}:
        raise HTTPException(status_code=400, detail="deck_mode must be one of: full, matching, custom")
    return mode


def format_percent(value: Any) -> str:
    if value is None or value == "":
        return MISSING_PPT_VALUE
    try:
        number = float(str(value).replace("%", "").replace(",", ""))
        if number <= 1:
            number *= 100
        return f"{number:.2f}%".rstrip("0").rstrip(".") + "%" if False else f"{number:.2f}%"
    except Exception:
        return safe_ppt(value)


def format_integer(value: Any) -> str:
    if value is None or value == "":
        return MISSING_PPT_VALUE
    try:
        return f"{int(round(float(str(value).replace(',', '')))):,}"
    except Exception:
        return safe_ppt(value)


def format_currency(value: Any, symbol: str = "£") -> str:
    if value is None or value == "":
        return MISSING_PPT_VALUE
    try:
        return f"{symbol}{float(str(value).replace(',', '').replace(symbol, '')):,.2f}"
    except Exception:
        return safe_ppt(value)


def file_response_payload(filename: str, content: bytes) -> Dict[str, Any]:
    if len(content) > MAX_RETURN_FILE_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"Generated PPT is {len(content)} bytes and exceeds MAX_RETURN_FILE_BYTES={MAX_RETURN_FILE_BYTES}.",
        )
    return {
        "filename": filename,
        "mime_type": PPTX_MIME,
        "data": base64.b64encode(content).decode("utf-8"),
        "size_bytes": len(content),
    }


def decode_uploaded_excel(req: Any) -> Tuple[bytes, str]:
    """Supports file_base64 and OpenAI action-style file refs with embedded base64 data.

    In production OpenAI Actions, file download may be proxied by the platform. This function is
    intentionally conservative: it supports embedded `data`/`file_base64` and returns a clear error
    if the action layer has not provided bytes.
    """
    if getattr(req, "file_base64", None):
        filename = getattr(req, "filename", None) or "EOC_Report.xlsx"
        return base64.b64decode(req.file_base64), filename

    refs = getattr(req, "openaiFileIdRefs", None) or []
    for ref in refs:
        if ref.data:
            filename = ref.name or getattr(req, "filename", None) or "EOC_Report.xlsx"
            return base64.b64decode(ref.data), filename

    raise HTTPException(
        status_code=400,
        detail="No Excel bytes supplied. Provide file_base64 or openaiFileIdRefs with embedded data.",
    )


# =============================================================================
# EOC validation / parsing
# =============================================================================

@dataclass
class ParsedEOC:
    display_values: Dict[str, str]
    mapped_values: Dict[str, str]
    raw_summary: Dict[str, Any]


def read_excel_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    try:
        with io.BytesIO(file_bytes) as stream:
            return pd.read_excel(stream, sheet_name=None, header=None, dtype=object)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Unable to read Excel file: {exc}") from exc


def dataframe_to_search_blob(df: pd.DataFrame, max_rows: int = 80, max_cols: int = 30) -> str:
    subset = df.iloc[:max_rows, :max_cols].copy()
    values: List[str] = []
    for row in subset.values.tolist():
        values.extend(clean_text(cell) for cell in row if clean_text(cell))
    return " | ".join(values)


def find_value_near_label(df: pd.DataFrame, label_aliases: Sequence[str], right_limit: int = 6, down_limit: int = 2) -> str:
    aliases = [a.lower() for a in label_aliases]
    rows, cols = df.shape
    for r in range(rows):
        for c in range(cols):
            cell = clean_text(df.iat[r, c]).lower()
            if not cell:
                continue
            if any(alias in cell for alias in aliases):
                for dc in range(1, right_limit + 1):
                    if c + dc < cols:
                        value = clean_text(df.iat[r, c + dc])
                        if value and not any(alias in value.lower() for alias in aliases):
                            return value
                for dr in range(1, down_limit + 1):
                    if r + dr < rows:
                        value = clean_text(df.iat[r + dr, c])
                        if value and not any(alias in value.lower() for alias in aliases):
                            return value
    return ""


def detect_header_row(df: pd.DataFrame, required_any: Sequence[str], max_scan_rows: int = 60) -> Optional[int]:
    required = [x.lower() for x in required_any]
    max_r = min(max_scan_rows, df.shape[0])
    for r in range(max_r):
        row_text = " | ".join(clean_text(v).lower() for v in df.iloc[r].tolist())
        hits = sum(1 for item in required if item in row_text)
        if hits >= 2:
            return r
    return None


def table_from_header(df: pd.DataFrame, header_row: int) -> pd.DataFrame:
    header = [clean_text(x) or f"Column_{i}" for i, x in enumerate(df.iloc[header_row].tolist())]
    data = df.iloc[header_row + 1 :].copy()
    data.columns = header
    data = data.dropna(how="all")
    return data


def get_column(table: pd.DataFrame, aliases: Sequence[str]) -> Optional[str]:
    alias_norm = [a.lower() for a in aliases]
    for col in table.columns:
        c = clean_text(col).lower()
        if any(a == c or a in c for a in alias_norm):
            return col
    return None


def coerce_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    text = clean_text(value)
    if not text:
        return None
    text = text.replace("%", "").replace(",", "").replace("£", "").replace("$", "").replace("€", "")
    try:
        return float(text)
    except Exception:
        return None


def extract_top_performers(table: pd.DataFrame, metric_aliases: Sequence[str], name_aliases: Sequence[str]) -> List[Dict[str, str]]:
    name_col = get_column(table, name_aliases)
    metric_col = get_column(table, metric_aliases)
    if not name_col or not metric_col:
        return []

    rows = []
    for _, row in table.iterrows():
        name = clean_text(row.get(name_col))
        metric_raw = row.get(metric_col)
        metric = coerce_number(metric_raw)
        if not name or name.lower() in {"total", "grand total"} or metric is None:
            continue
        rows.append((name, metric, metric_raw))

    rows.sort(key=lambda x: x[1], reverse=True)
    return [
        {"name": name, "value": format_percent(metric_raw)}
        for name, _, metric_raw in rows[:3]
    ]


def parse_eoc(file_bytes: bytes) -> ParsedEOC:
    sheets = read_excel_sheets(file_bytes)
    all_blob = " | ".join(dataframe_to_search_blob(df) for df in sheets.values())

    campaign = ""
    client = ""
    markets = ""
    formats = ""
    live_dates = ""
    spend = ""
    impressions = ""
    ctr = ""
    engagement_rate = ""
    vcr = ""
    on_screen = ""

    candidate_tables: List[pd.DataFrame] = []

    for _, df in sheets.items():
        campaign = campaign or find_value_near_label(df, ["campaign", "campaign name"])
        client = client or find_value_near_label(df, ["client", "advertiser", "brand"])
        markets = markets or find_value_near_label(df, ["geo", "market", "markets", "country"])
        formats = formats or find_value_near_label(df, ["format", "formats", "product"])
        live_dates = live_dates or find_value_near_label(df, ["live dates", "dates", "start date", "end date"])
        spend = spend or find_value_near_label(df, ["spend", "budget", "campaign budget"])
        impressions = impressions or find_value_near_label(df, ["delivered impressions", "impressions", "delivered overall"])
        ctr = ctr or find_value_near_label(df, ["ctr", "click through rate"])
        engagement_rate = engagement_rate or find_value_near_label(df, ["engagement rate", "engagement"])
        vcr = vcr or find_value_near_label(df, ["video completion rate", "vcr"])
        on_screen = on_screen or find_value_near_label(df, ["mobkoi on screen", "on screen", "viewability", "mrc viewability"])

        header_row = detect_header_row(
            df,
            required_any=["site", "publisher", "title", "geo", "market", "ctr", "engagement", "vcr", "impressions"],
        )
        if header_row is not None:
            table = table_from_header(df, header_row)
            if len(table.columns) >= 3:
                candidate_tables.append(table)

    top_ctr: List[Dict[str, str]] = []
    top_engagement: List[Dict[str, str]] = []
    top_vcr: List[Dict[str, str]] = []

    for table in candidate_tables:
        name_aliases = ["site", "publisher", "title", "property", "placement", "app", "website"]
        top_ctr = top_ctr or extract_top_performers(table, ["ctr", "click through rate"], name_aliases)
        top_engagement = top_engagement or extract_top_performers(table, ["engagement rate", "engagement"], name_aliases)
        top_vcr = top_vcr or extract_top_performers(table, ["video completion rate", "vcr"], name_aliases)

    # Fallback campaign name from workbook blob if no explicit label is found.
    if not campaign:
        campaign_match = re.search(r"([A-Z][A-Za-z0-9 &'-]{2,80})", all_blob)
        campaign = campaign_match.group(1) if campaign_match else "Campaign"

    mapped_values = {
        "{{CAMPAIGN_NAME}}": safe_ppt(campaign),
        "{{CLIENT_NAME}}": safe_ppt(client),
        "{{LIVE_DATES_SHORT}}": safe_ppt(live_dates),
        "{{LIVE_DATES_FULL}}": safe_ppt(live_dates),
        "{{MARKETS}}": safe_ppt(markets),
        "{{FORMATS}}": safe_ppt(formats),
        "{{CAMPAIGN_BUDGET}}": format_currency(spend) if spend else MISSING_PPT_VALUE,
        "{{DELIVERED_IMPRESSIONS}}": format_integer(impressions) if impressions else MISSING_PPT_VALUE,
        "{{PERFORMANCE_CTR}}": format_percent(ctr) if ctr else MISSING_PPT_VALUE,
        "{{PERFORMANCE_ENGAGEMENT_RATE}}": format_percent(engagement_rate) if engagement_rate else MISSING_PPT_VALUE,
        "{{PERFORMANCE_VCR}}": format_percent(vcr) if vcr else MISSING_PPT_VALUE,
        "{{PERFORMANCE_VIEWABILITY}}": format_percent(on_screen) if on_screen else MISSING_PPT_VALUE,
        "{{TOP_CTR_1}}": top_ctr[0]["name"] if len(top_ctr) > 0 else MISSING_PPT_VALUE,
        "{{TOP_CTR_1_VALUE}}": top_ctr[0]["value"] if len(top_ctr) > 0 else MISSING_PPT_VALUE,
        "{{TOP_CTR_2}}": top_ctr[1]["name"] if len(top_ctr) > 1 else MISSING_PPT_VALUE,
        "{{TOP_CTR_2_VALUE}}": top_ctr[1]["value"] if len(top_ctr) > 1 else MISSING_PPT_VALUE,
        "{{TOP_CTR_3}}": top_ctr[2]["name"] if len(top_ctr) > 2 else MISSING_PPT_VALUE,
        "{{TOP_CTR_3_VALUE}}": top_ctr[2]["value"] if len(top_ctr) > 2 else MISSING_PPT_VALUE,
        "{{TOP_ENGAGEMENT_1}}": top_engagement[0]["name"] if len(top_engagement) > 0 else MISSING_PPT_VALUE,
        "{{TOP_ENGAGEMENT_1_VALUE}}": top_engagement[0]["value"] if len(top_engagement) > 0 else MISSING_PPT_VALUE,
        "{{TOP_VCR_1}}": top_vcr[0]["name"] if len(top_vcr) > 0 else MISSING_PPT_VALUE,
        "{{TOP_VCR_1_VALUE}}": top_vcr[0]["value"] if len(top_vcr) > 0 else MISSING_PPT_VALUE,
    }

    display_values = {
        "Campaign": safe_display(campaign),
        "Client": safe_display(client),
        "Live Dates": safe_display(live_dates),
        "Markets": safe_display(markets),
        "Formats": safe_display(formats),
        "Campaign Budget": mapped_values["{{CAMPAIGN_BUDGET}}"],
        "Delivered Impressions": mapped_values["{{DELIVERED_IMPRESSIONS}}"],
        "CTR": mapped_values["{{PERFORMANCE_CTR}}"],
        "Engagement Rate": mapped_values["{{PERFORMANCE_ENGAGEMENT_RATE}}"],
        "VCR": mapped_values["{{PERFORMANCE_VCR}}"],
        "On Screen": mapped_values["{{PERFORMANCE_VIEWABILITY}}"],
    }

    raw_summary = {
        "top_ctr": top_ctr,
        "top_engagement": top_engagement,
        "top_vcr": top_vcr,
        "sheets": list(sheets.keys()),
    }

    return ParsedEOC(display_values=display_values, mapped_values=mapped_values, raw_summary=raw_summary)


def build_exec_summary(parsed: ParsedEOC) -> str:
    campaign = parsed.display_values.get("Campaign", MISSING_DISPLAY_VALUE)
    markets = parsed.display_values.get("Markets", MISSING_DISPLAY_VALUE)
    impressions = parsed.display_values.get("Delivered Impressions", MISSING_DISPLAY_VALUE)
    ctr = parsed.display_values.get("CTR", MISSING_DISPLAY_VALUE)
    engagement = parsed.display_values.get("Engagement Rate", MISSING_DISPLAY_VALUE)
    vcr = parsed.display_values.get("VCR", MISSING_DISPLAY_VALUE)

    return (
        f"{campaign} delivered a premium mobile campaign across {markets}, generating {impressions} delivered impressions. "
        f"Performance was led by CTR at {ctr}, engagement rate at {engagement}, and video completion at {vcr}, "
        "with the strongest results coming from the top-performing environments identified in the EOC report."
    )


# =============================================================================
# PowerPoint replacement helpers
# =============================================================================

def iter_shapes_recursive(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes_recursive(shape.shapes)


def replace_text_preserve_runs(shape, replacements: Dict[str, str]) -> None:
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        return

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            text = run.text
            if not text:
                continue
            new_text = text
            for placeholder, value in replacements.items():
                if placeholder in new_text:
                    new_text = new_text.replace(placeholder, value)
            if new_text != text:
                run.text = new_text


def replace_text_in_presentation(prs: Presentation, replacements: Dict[str, str]) -> None:
    for slide in prs.slides:
        for shape in iter_shapes_recursive(slide.shapes):
            replace_text_preserve_runs(shape, replacements)


def save_presentation_to_bytes(prs: Presentation) -> bytes:
    stream = io.BytesIO()
    prs.save(stream)
    return stream.getvalue()


# =============================================================================
# SECTION_ID helpers for slide decks
# =============================================================================

def shape_text(shape) -> str:
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        return ""
    return clean_text(shape.text_frame.text)


def get_slide_section_ids(slide) -> List[str]:
    ids: List[str] = []
    for shape in iter_shapes_recursive(slide.shapes):
        text = shape_text(shape)
        if not text:
            continue
        # Supports exact marker boxes and inline markers.
        for match in re.finditer(r"SECTION_ID\s*:\s*([A-Za-z0-9_ -]+)", text, flags=re.IGNORECASE):
            sid = canonical_section_id(match.group(1))
            if sid and sid not in ids:
                ids.append(sid)
    return ids


def is_section_marker_shape(shape) -> bool:
    text = shape_text(shape)
    return bool(re.search(r"^\s*SECTION_ID\s*:", text, flags=re.IGNORECASE))


def remove_shape_from_slide(slide, shape) -> None:
    try:
        element = shape._element
        element.getparent().remove(element)
    except Exception:
        # Best-effort cleanup. Never fail PPT generation because of a marker shape.
        pass


def remove_section_id_text_boxes(prs: Presentation) -> None:
    for slide in prs.slides:
        shapes_to_remove = [shape for shape in iter_shapes_recursive(slide.shapes) if is_section_marker_shape(shape)]
        for shape in shapes_to_remove:
            remove_shape_from_slide(slide, shape)


def delete_slides_by_index(prs: Presentation, indexes_to_delete: Sequence[int]) -> None:
    """Delete slides from a python-pptx Presentation.

    Uses internal slide id list because python-pptx has no public delete API.
    Delete in reverse order to keep indexes stable.
    """
    slide_id_list = prs.slides._sldIdLst
    for index in sorted(indexes_to_delete, reverse=True):
        if index < 0 or index >= len(prs.slides):
            continue
        slide_id = slide_id_list[index]
        rel_id = slide_id.rId
        prs.part.drop_rel(rel_id)
        slide_id_list.remove(slide_id)


def filter_slide_deck_by_sections(prs: Presentation, keep_sections: List[str]) -> Dict[str, Any]:
    keep_set = set(keep_sections)
    indexes_to_delete: List[int] = []
    slide_report: List[Dict[str, Any]] = []

    for idx, slide in enumerate(prs.slides):
        slide_ids = get_slide_section_ids(slide)

        # Slides without SECTION_ID are treated as global/admin slides and retained.
        keep = True if not slide_ids else bool(keep_set.intersection(slide_ids))

        slide_report.append(
            {
                "slide_index": idx + 1,
                "section_ids": slide_ids,
                "kept": keep,
            }
        )

        if not keep:
            indexes_to_delete.append(idx)

    delete_slides_by_index(prs, indexes_to_delete)

    return {
        "requested_sections": keep_sections,
        "deleted_slide_count": len(indexes_to_delete),
        "kept_slide_count": len(prs.slides),
        "slide_report": slide_report,
    }


def resolve_deck_sections(
    deck_mode: str,
    one_pager_sections: Optional[Sequence[str]],
    section_selections: Optional[Sequence[str]],
    custom_deck_sections: Optional[Sequence[str]],
) -> Optional[List[str]]:
    mode = normalise_deck_mode(deck_mode)

    if mode == "full":
        return None

    if mode == "matching":
        source = one_pager_sections or section_selections or []
        return canonical_section_list(source, include_base=True)

    # custom
    source = custom_deck_sections or []
    if not source:
        raise HTTPException(status_code=400, detail="custom_deck_sections is required when deck_mode is custom.")
    return canonical_section_list(source, include_base=True)


# =============================================================================
# One Pager assembly
# =============================================================================

def generate_one_pager_pptx(
    parsed: ParsedEOC,
    exec_summary: Optional[str],
    section_selections: Optional[Sequence[str]],
) -> Tuple[bytes, Dict[str, Any]]:
    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail=f"One Pager template not found: {TEMPLATE_PATH}")

    prs = Presentation(str(TEMPLATE_PATH))
    sections = canonical_section_list(section_selections, include_base=True)

    replacements = dict(parsed.mapped_values)
    replacements["{{EXEC_SUMMARY}}"] = safe_ppt(exec_summary or build_exec_summary(parsed))

    replace_text_in_presentation(prs, replacements)

    # Stable workflow preservation:
    # This endpoint intentionally does not delete or restructure one-pager content.
    # Dynamic section stacking/copying can remain in the existing stable implementation.
    # The new deck filtering logic is isolated to slide-deck functions below.

    content = save_presentation_to_bytes(prs)
    meta = {
        "sections": sections,
        "template": str(TEMPLATE_PATH),
        "slide_count": len(prs.slides),
    }
    return content, meta


# =============================================================================
# Slide Deck generation
# =============================================================================

def generate_slide_deck_pptx(
    parsed: ParsedEOC,
    exec_summary: Optional[str],
    deck_mode: str,
    section_selections: Optional[Sequence[str]],
    one_pager_sections: Optional[Sequence[str]],
    custom_deck_sections: Optional[Sequence[str]],
) -> Tuple[bytes, Dict[str, Any]]:
    if not SLIDE_DECK_TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail=f"Slide Deck template not found: {SLIDE_DECK_TEMPLATE_PATH}")

    mode = normalise_deck_mode(deck_mode)
    prs = Presentation(str(SLIDE_DECK_TEMPLATE_PATH))

    replacements = dict(parsed.mapped_values)
    replacements["{{EXEC_SUMMARY}}"] = safe_ppt(exec_summary or build_exec_summary(parsed))

    filter_meta: Dict[str, Any] = {
        "deck_mode": mode,
        "template": str(SLIDE_DECK_TEMPLATE_PATH),
        "initial_slide_count": len(prs.slides),
    }

    deck_sections = resolve_deck_sections(
        deck_mode=mode,
        one_pager_sections=one_pager_sections,
        section_selections=section_selections,
        custom_deck_sections=custom_deck_sections,
    )

    if deck_sections is not None:
        filter_meta.update(filter_slide_deck_by_sections(prs, deck_sections))
    else:
        filter_meta.update(
            {
                "requested_sections": "ALL",
                "deleted_slide_count": 0,
                "kept_slide_count": len(prs.slides),
                "slide_report": [
                    {
                        "slide_index": idx + 1,
                        "section_ids": get_slide_section_ids(slide),
                        "kept": True,
                    }
                    for idx, slide in enumerate(prs.slides)
                ],
            }
        )

    # Remove marker text boxes after filtering so detection still works before deletion.
    remove_section_id_text_boxes(prs)

    # Replace placeholders after deleting slides and marker boxes.
    replace_text_in_presentation(prs, replacements)

    filter_meta["final_slide_count"] = len(prs.slides)
    content = save_presentation_to_bytes(prs)
    return content, filter_meta


# =============================================================================
# API endpoints
# =============================================================================

@app.get("/health")
def health() -> Dict[str, Any]:
    return {
        "status": "ok",
        "version": APP_VERSION,
        "template_found": TEMPLATE_PATH.exists(),
        "slide_deck_template_found": SLIDE_DECK_TEMPLATE_PATH.exists(),
        "rules_master_found": RULES_MASTER_PATH.exists(),
        "template_path": str(TEMPLATE_PATH),
        "slide_deck_template_path": str(SLIDE_DECK_TEMPLATE_PATH),
        "rules_master_path": str(RULES_MASTER_PATH),
        "supported_deck_modes": ["full", "matching", "custom"],
    }


@app.post("/validate-eoc")
def validate_eoc(req: ValidateEocRequest) -> Dict[str, Any]:
    file_bytes, filename = decode_uploaded_excel(req)
    parsed = parse_eoc(file_bytes)
    exec_summary = build_exec_summary(parsed)

    return {
        "status": "validated",
        "version": APP_VERSION,
        "filename": filename,
        "display_values": parsed.display_values,
        "mapped_values": parsed.mapped_values,
        "top_performers": {
            "ctr": parsed.raw_summary.get("top_ctr", []),
            "engagement": parsed.raw_summary.get("top_engagement", []),
            "vcr": parsed.raw_summary.get("top_vcr", []),
        },
        "exec_summary": exec_summary,
        "missing_value_display": MISSING_DISPLAY_VALUE,
        "missing_value_ppt": MISSING_PPT_VALUE,
    }


@app.post("/generate-exec-summary")
def generate_exec_summary(req: GenerateExecSummaryRequest) -> Dict[str, Any]:
    file_bytes, filename = decode_uploaded_excel(req)
    parsed = parse_eoc(file_bytes)

    pptx_bytes, meta = generate_one_pager_pptx(
        parsed=parsed,
        exec_summary=req.exec_summary,
        section_selections=req.section_selections or req.one_pager_sections,
    )

    campaign = slug_filename(parsed.mapped_values.get("{{CAMPAIGN_NAME}}", "Campaign"))
    output_name = req.output_filename or f"PCA One Pager_{campaign}_{now_stamp()}.pptx"

    return {
        "status": "generated",
        "version": APP_VERSION,
        "source_filename": filename,
        "type": "one_pager",
        "metadata": meta,
        "file": file_response_payload(output_name, pptx_bytes),
    }


@app.post("/generate-modular-one-pager")
def generate_modular_one_pager(req: GenerateExecSummaryRequest) -> Dict[str, Any]:
    # Alias endpoint to preserve the modular GPT workflow name.
    return generate_exec_summary(req)


@app.post("/generate-slide-deck")
def generate_slide_deck(req: GenerateSlideDeckRequest) -> Dict[str, Any]:
    file_bytes, filename = decode_uploaded_excel(req)
    parsed = parse_eoc(file_bytes)

    pptx_bytes, meta = generate_slide_deck_pptx(
        parsed=parsed,
        exec_summary=req.exec_summary,
        deck_mode=req.deck_mode,
        section_selections=req.section_selections,
        one_pager_sections=req.one_pager_sections,
        custom_deck_sections=req.custom_deck_sections,
    )

    campaign = slug_filename(parsed.mapped_values.get("{{CAMPAIGN_NAME}}", "Campaign"))
    output_name = req.output_filename or f"PCA Slides_{campaign}_{now_stamp()}.pptx"

    return {
        "status": "generated",
        "version": APP_VERSION,
        "source_filename": filename,
        "type": "slide_deck",
        "deck_mode": normalise_deck_mode(req.deck_mode),
        "metadata": meta,
        "file": file_response_payload(output_name, pptx_bytes),
    }


@app.post("/generate-one-pager-and-slides")
def generate_one_pager_and_slides(req: GenerateOnePagerAndSlidesRequest) -> Dict[str, Any]:
    file_bytes, filename = decode_uploaded_excel(req)
    parsed = parse_eoc(file_bytes)

    selected_sections = req.one_pager_sections or req.section_selections or []

    one_pager_bytes, one_pager_meta = generate_one_pager_pptx(
        parsed=parsed,
        exec_summary=req.exec_summary,
        section_selections=selected_sections,
    )

    slide_deck_bytes, slide_deck_meta = generate_slide_deck_pptx(
        parsed=parsed,
        exec_summary=req.exec_summary,
        deck_mode=req.deck_mode,
        section_selections=selected_sections,
        one_pager_sections=selected_sections,
        custom_deck_sections=req.custom_deck_sections,
    )

    campaign = slug_filename(parsed.mapped_values.get("{{CAMPAIGN_NAME}}", "Campaign"))
    one_pager_name = req.one_pager_filename or f"PCA One Pager_{campaign}_{now_stamp()}.pptx"
    slide_deck_name = req.slide_deck_filename or f"PCA Slides_{campaign}_{now_stamp()}.pptx"

    return {
        "status": "generated",
        "version": APP_VERSION,
        "source_filename": filename,
        "types": ["one_pager", "slide_deck"],
        "deck_mode": normalise_deck_mode(req.deck_mode),
        "one_pager_metadata": one_pager_meta,
        "slide_deck_metadata": slide_deck_meta,
        "files": [
            file_response_payload(one_pager_name, one_pager_bytes),
            file_response_payload(slide_deck_name, slide_deck_bytes),
        ],
    }


# =============================================================================
# Local run
# =============================================================================

if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=int(os.getenv("PORT", "8000")), reload=True)

