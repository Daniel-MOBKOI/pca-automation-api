import base64
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import httpx
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from pptx import Presentation


APP_VERSION = "10.0.5-hybrid-production"

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = Path(
    os.getenv(
        "TEMPLATE_PATH",
        BASE_DIR / "templates" / "exec_summary_master.pptx",
    )
)

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MAX_RETURN_FILE_BYTES = 10 * 1024 * 1024


app = FastAPI(
    title="PCA Automation API",
    version=APP_VERSION,
    description="Validates EOC Excel files and generates editable Exec Summary PPT decks.",
)


# -----------------------------
# OpenAI Actions file payload
# -----------------------------

class OpenAIFileRef(BaseModel):
    id: Optional[str] = None
    name: Optional[str] = None
    mime_type: Optional[str] = None
    download_link: Optional[str] = None


class FileRefsPayload(BaseModel):
    openaiFileIdRefs: List[OpenAIFileRef] = Field(
        ...,
        description="Files supplied by OpenAI Actions. Upload exactly one EOC Excel file.",
    )


# -----------------------------
# Formatting helpers
# -----------------------------

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def normalise(value: Any) -> str:
    value = clean_text(value).lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    return value.strip("_")


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    text = clean_text(value)
    if not text:
        return None

    text = text.replace(",", "").replace("£", "").replace("$", "").replace("€", "").strip()

    is_percent = "%" in text
    text = text.replace("%", "")

    try:
        num = float(text)
        if is_percent:
            return num
        return num
    except Exception:
        return None


def fmt_int(value: Any) -> str:
    num = to_float(value)
    if num is None:
        return ""
    return f"{int(round(num)):,}"


def fmt_currency(value: Any, currency: str = "") -> str:
    num = to_float(value)
    if num is None:
        return ""

    symbol = ""
    cur = clean_text(currency).upper()

    if cur in ["GBP", "POUND", "POUNDS", "£"]:
        symbol = "£"
    elif cur in ["EUR", "EURO", "EUROS", "€"]:
        symbol = "€"
    elif cur in ["USD", "DOLLAR", "DOLLARS", "$"]:
        symbol = "$"

    return f"{symbol}{num:,.0f}"


def fmt_percent(value: Any, decimals: int = 2) -> str:
    num = to_float(value)
    if num is None:
        return ""

    # If Excel stores 0.0032 for 0.32%, convert to 0.32
    if 0 < num <= 1:
        num *= 100

    return f"{num:.{decimals}f}%"


def safe_ctr(clicks: Any, impressions: Any) -> str:
    clk = to_float(clicks)
    imp = to_float(impressions)

    if clk is None or imp is None or imp <= 0:
        return ""

    return f"{(clk / imp) * 100:.2f}%"


def parse_date(value: Any) -> Optional[datetime]:
    if value is None or clean_text(value) == "":
        return None

    try:
        dt = pd.to_datetime(value, errors="coerce")
        if pd.isna(dt):
            return None

        parsed = dt.to_pydatetime()

        # Avoid bogus Excel/null fallback dates.
        if parsed.year < 2000:
            return None

        return parsed
    except Exception:
        return None


def format_date_short(start: Optional[datetime], end: Optional[datetime]) -> str:
    if not start and not end:
        return ""
    if start and not end:
        return f"{start.day} {start.strftime('%b')}"
    if end and not start:
        return f"{end.day} {end.strftime('%b')}"

    if start.year == end.year:
        if start.month == end.month:
            return f"{start.day} - {end.day} {end.strftime('%b')}"
        return f"{start.day} {start.strftime('%b')} - {end.day} {end.strftime('%b')}"

    return f"{start.day} {start.strftime('%b %Y')} - {end.day} {end.strftime('%b %Y')}"


def format_date_full(start: Optional[datetime], end: Optional[datetime]) -> str:
    if not start and not end:
        return ""
    if start and not end:
        return f"{start.day} {start.strftime('%B %Y')}"
    if end and not start:
        return f"{end.day} {end.strftime('%B %Y')}"

    if start.year == end.year:
        if start.month == end.month:
            return f"{start.day} - {end.day} {end.strftime('%B %Y')}"
        return f"{start.day} {start.strftime('%B')} - {end.day} {end.strftime('%B %Y')}"

    return f"{start.day} {start.strftime('%B %Y')} - {end.day} {end.strftime('%B %Y')}"


# -----------------------------
# File download
# -----------------------------

async def download_openai_file(file_ref: OpenAIFileRef, dest_dir: Path) -> Path:
    if not file_ref.download_link:
        raise HTTPException(
            status_code=400,
            detail="Missing download_link in openaiFileIdRefs.",
        )

    safe_name = file_ref.name or "uploaded_eoc.xlsx"
    safe_name = re.sub(r"[^A-Za-z0-9._ -]+", "_", safe_name)
    output_path = dest_dir / safe_name

    async with httpx.AsyncClient(timeout=60.0, follow_redirects=True) as client:
        response = await client.get(file_ref.download_link)
        response.raise_for_status()
        output_path.write_bytes(response.content)

    return output_path


def read_eoc(file_path: Path) -> Dict[str, pd.DataFrame]:
    try:
        return pd.read_excel(file_path, sheet_name=None, header=None, engine="openpyxl")
    except Exception as exc:
        raise HTTPException(
            status_code=400,
            detail=f"Could not read EOC Excel file: {exc}",
        )


# -----------------------------
# Hybrid parser
# -----------------------------

def get_cell(df: pd.DataFrame, r: int, c: int) -> Any:
    if r < 0 or c < 0:
        return ""
    if r >= df.shape[0] or c >= df.shape[1]:
        return ""
    return df.iat[r, c]


def find_label_value(
    sheets: Dict[str, pd.DataFrame],
    labels: List[str],
    right_offsets: Tuple[int, ...] = (1, 2),
    below_offsets: Tuple[int, ...] = (1,),
) -> Any:
    """
    Primary label lookup.
    Searches for labels and returns nearby values to the right or below.
    This is used only when labels are clearly found.
    """
    wanted = [normalise(x) for x in labels]

    for df in sheets.values():
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_norm = normalise(get_cell(df, r, c))
                if not cell_norm:
                    continue

                matched = False

                for label in wanted:
                    if cell_norm == label or label in cell_norm:
                        matched = True
                        break

                if not matched:
                    continue

                for off in right_offsets:
                    val = get_cell(df, r, c + off)
                    if clean_text(val):
                        return val

                for off in below_offsets:
                    val = get_cell(df, r + off, c)
                    if clean_text(val):
                        return val

    return ""


def find_campaign_row_values(sheets: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    """
    Detects common EOC summary blocks like:
    Campaign | Currency | Sold Paid Units
    Prada... | Euro     | 1136364

    Campaign | Impressions | Clicks
    Prada... | 1175313     | 9153
    """
    result: Dict[str, Any] = {}

    for df in sheets.values():
        for r in range(df.shape[0] - 1):
            row_norm = [normalise(get_cell(df, r, c)) for c in range(df.shape[1])]

            if "campaign" not in row_norm:
                continue

            campaign_col = row_norm.index("campaign")
            campaign_value = get_cell(df, r + 1, campaign_col)

            if clean_text(campaign_value):
                result.setdefault("CAMPAIGN_NAME", campaign_value)

            for c, header in enumerate(row_norm):
                value = get_cell(df, r + 1, c)

                if not clean_text(value):
                    continue

                if header in ["currency"]:
                    result["CURRENCY"] = value
                elif header in ["market", "markets", "country", "countries"]:
                    result["MARKETS"] = value
                elif header in ["sold_paid_units", "booked_impressions", "planned_impressions"]:
                    result["SOLD_PAID_UNITS"] = value
                elif header in ["impressions", "delivered_impressions"]:
                    result["IMPRESSIONS"] = value
                elif header in ["clicks", "total_clicks"]:
                    result["CLICKS"] = value
                elif header in ["ctr", "click_through_rate"]:
                    result["CTR"] = value
                elif header in ["video_starts", "starts"]:
                    result["VIDEO_STARTS"] = value
                elif header in ["video_completes", "completed_views", "completions"]:
                    result["COMPLETED_VIEWS"] = value
                elif header in ["vcr", "video_completion_rate"]:
                    result["VCR"] = value
                elif header in ["engagements", "total_engagements"]:
                    result["ENGAGEMENTS"] = value
                elif header in ["engagement_rate", "er"]:
                    result["ENGAGEMENT_RATE"] = value
                elif header in ["viewability", "viewability_rate"]:
                    result["VIEWABILITY"] = value
                elif header in ["spend", "budget", "media_spend"]:
                    result["BUDGET"] = value

    return result


def extract_metric_table(sheets: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    """
    Finds metric-style tables where metrics are headers or row labels.
    Used after campaign row extraction.
    """
    result: Dict[str, Any] = {}

    metric_aliases = {
        "IMPRESSIONS": ["impressions", "delivered_impressions"],
        "CLICKS": ["clicks", "total_clicks"],
        "CTR": ["ctr", "click_through_rate"],
        "ENGAGEMENTS": ["engagements", "total_engagements"],
        "ENGAGEMENT_RATE": ["engagement_rate", "er"],
        "VIEWABILITY": ["viewability", "viewability_rate"],
        "VCR": ["vcr", "video_completion_rate", "completion_rate"],
        "COMPLETED_VIEWS": ["completed_views", "video_completes", "completions"],
        "VIDEO_STARTS": ["video_starts", "starts"],
        "BUDGET": ["budget", "spend", "media_spend"],
    }

    for df in sheets.values():
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell = normalise(get_cell(df, r, c))
                if not cell:
                    continue

                for canonical, aliases in metric_aliases.items():
                    if cell in aliases or any(alias in cell for alias in aliases):
                        # Prefer right cell, then below.
                        right = get_cell(df, r, c + 1)
                        below = get_cell(df, r + 1, c)

                        if clean_text(right):
                            result.setdefault(canonical, right)
                        elif clean_text(below):
                            result.setdefault(canonical, below)

    return result


def extract_market_from_sheet(sheets: Dict[str, pd.DataFrame]) -> str:
    val = find_label_value(
        sheets,
        ["Market", "Markets", "Country", "Countries"],
        right_offsets=(1, 2),
        below_offsets=(1,),
    )
    return clean_text(val)


def extract_dates(sheets: Dict[str, pd.DataFrame]) -> Tuple[Optional[datetime], Optional[datetime]]:
    start_raw = find_label_value(
        sheets,
        ["Start Date", "Live Start", "Campaign Start", "Flight Start", "From"],
    )
    end_raw = find_label_value(
        sheets,
        ["End Date", "Live End", "Campaign End", "Flight End", "To"],
    )

    start = parse_date(start_raw)
    end = parse_date(end_raw)

    # Fallback: scan plausible date cells.
    if not start or not end:
        dates: List[datetime] = []

        for df in sheets.values():
            for r in range(df.shape[0]):
                for c in range(df.shape[1]):
                    dt = parse_date(get_cell(df, r, c))
                    if dt:
                        dates.append(dt)

        if dates:
            dates = sorted(dates)
            start = start or dates[0]
            end = end or dates[-1]

    return start, end


def extract_top_titles(sheets: Dict[str, pd.DataFrame], metric: str, limit: int = 3) -> List[Dict[str, str]]:
    """
    Finds rows with title/site and metric columns.
    Prevents wrong CTRs by treating percent values safely.
    """
    name_headers = ["site", "publisher", "title", "placement", "domain", "inventory"]
    metric_headers = {
        "CTR": ["ctr", "click_through_rate"],
        "ENGAGEMENT_RATE": ["engagement_rate", "er"],
        "VCR": ["vcr", "video_completion_rate", "completion_rate"],
    }.get(metric, [])

    found: List[Tuple[float, str, Any]] = []

    for df in sheets.values():
        for r in range(df.shape[0]):
            headers = [normalise(get_cell(df, r, c)) for c in range(df.shape[1])]

            name_col = None
            metric_col = None

            for c, h in enumerate(headers):
                if h in name_headers and name_col is None:
                    name_col = c
                if h in metric_headers and metric_col is None:
                    metric_col = c

            if name_col is None or metric_col is None:
                continue

            for rr in range(r + 1, df.shape[0]):
                name = clean_text(get_cell(df, rr, name_col))
                raw_metric = get_cell(df, rr, metric_col)
                metric_num = to_float(raw_metric)

                if not name or metric_num is None:
                    continue

                # If stored as decimal, convert for sorting/display later.
                sort_num = metric_num * 100 if 0 < metric_num <= 1 else metric_num

                if sort_num < 0 or sort_num > 100:
                    # Avoid nonsense like 31,122% unless this is genuinely not a rate.
                    continue

                found.append((sort_num, name, raw_metric))

    found.sort(key=lambda x: x[0], reverse=True)

    return [
        {
            "name": item[1],
            "value": fmt_percent(item[2]),
        }
        for item in found[:limit]
    ]


def build_mapped_values(file_path: Path) -> Dict[str, str]:
    sheets = read_eoc(file_path)

    campaign_rows = find_campaign_row_values(sheets)
    metric_table = extract_metric_table(sheets)

    campaign_name = (
        clean_text(campaign_rows.get("CAMPAIGN_NAME"))
        or clean_text(find_label_value(sheets, ["Campaign Name", "Campaign"]))
    )

    client_name = clean_text(
        find_label_value(sheets, ["Client", "Advertiser", "Brand"])
    )

    markets = (
        clean_text(campaign_rows.get("MARKETS"))
        or extract_market_from_sheet(sheets)
    )

    currency = clean_text(campaign_rows.get("CURRENCY"))

    start_date, end_date = extract_dates(sheets)

    impressions = (
        campaign_rows.get("IMPRESSIONS")
        or metric_table.get("IMPRESSIONS")
        or find_label_value(sheets, ["Impressions", "Delivered Impressions"])
    )

    clicks = (
        campaign_rows.get("CLICKS")
        or metric_table.get("CLICKS")
        or find_label_value(sheets, ["Clicks", "Total Clicks"])
    )

    ctr = (
        campaign_rows.get("CTR")
        or metric_table.get("CTR")
        or find_label_value(sheets, ["CTR", "Click Through Rate", "Click-Through Rate"])
    )

    engagements = (
        campaign_rows.get("ENGAGEMENTS")
        or metric_table.get("ENGAGEMENTS")
        or find_label_value(sheets, ["Engagements", "Total Engagements"])
    )

    engagement_rate = (
        campaign_rows.get("ENGAGEMENT_RATE")
        or metric_table.get("ENGAGEMENT_RATE")
        or find_label_value(sheets, ["Engagement Rate", "ER"])
    )

    viewability = (
        campaign_rows.get("VIEWABILITY")
        or metric_table.get("VIEWABILITY")
        or find_label_value(sheets, ["Viewability", "Viewability Rate"])
    )

    completed_views = (
        campaign_rows.get("COMPLETED_VIEWS")
        or metric_table.get("COMPLETED_VIEWS")
        or find_label_value(sheets, ["Completed Views", "Video Completes", "Completions"])
    )

    video_starts = (
        campaign_rows.get("VIDEO_STARTS")
        or metric_table.get("VIDEO_STARTS")
        or find_label_value(sheets, ["Video Starts", "Starts"])
    )

    vcr = (
        campaign_rows.get("VCR")
        or metric_table.get("VCR")
        or find_label_value(sheets, ["VCR", "Video Completion Rate", "Completion Rate"])
    )

    budget = (
        campaign_rows.get("BUDGET")
        or metric_table.get("BUDGET")
        or find_label_value(sheets, ["Budget", "Spend", "Media Spend"])
    )

    sold_paid_units = campaign_rows.get("SOLD_PAID_UNITS")

    # Safe derived metrics.
    ctr_final = fmt_percent(ctr) if clean_text(ctr) else safe_ctr(clicks, impressions)

    if not clean_text(vcr) and to_float(video_starts) and to_float(completed_views):
        starts = to_float(video_starts)
        completes = to_float(completed_views)
        if starts and starts > 0 and completes is not None:
            vcr_final = f"{(completes / starts) * 100:.2f}%"
        else:
            vcr_final = ""
    else:
        vcr_final = fmt_percent(vcr)

    top_ctr = extract_top_titles(sheets, "CTR", 3)
    top_er = extract_top_titles(sheets, "ENGAGEMENT_RATE", 3)
    top_vcr = extract_top_titles(sheets, "VCR", 3)

    mapped = {
        "CAMPAIGN_NAME": campaign_name,
        "CLIENT_NAME": client_name,
        "MARKETS": markets,
        "FORMAT": clean_text(find_label_value(sheets, ["Format", "Product", "Creative Format"])),
        "LIVE_DATES_SHORT": format_date_short(start_date, end_date),
        "LIVE_DATES_FULL": format_date_full(start_date, end_date),

        "IMPRESSIONS": fmt_int(impressions),
        "CLICKS": fmt_int(clicks),
        "CTR": ctr_final,

        "ENGAGEMENTS": fmt_int(engagements),
        "ENGAGEMENT_RATE": fmt_percent(engagement_rate),

        "VIEWABILITY": fmt_percent(viewability),

        "VIDEO_STARTS": fmt_int(video_starts),
        "COMPLETED_VIEWS": fmt_int(completed_views),
        "VCR": vcr_final,
        "VIDEO_COMPLETION_RATE": vcr_final,

        "BUDGET": fmt_currency(budget, currency),
        "SOLD_PAID_UNITS": fmt_int(sold_paid_units),
        "CURRENCY": currency,
    }

    for idx in range(1, 4):
        ctr_row = top_ctr[idx - 1] if len(top_ctr) >= idx else {"name": "", "value": ""}
        er_row = top_er[idx - 1] if len(top_er) >= idx else {"name": "", "value": ""}
        vcr_row = top_vcr[idx - 1] if len(top_vcr) >= idx else {"name": "", "value": ""}

        mapped[f"TOP_TITLES_CTR_{idx}_NAME"] = ctr_row["name"]
        mapped[f"TOP_TITLES_CTR_{idx}_VALUE"] = ctr_row["value"]

        mapped[f"TOP_TITLES_ER_{idx}_NAME"] = er_row["name"]
        mapped[f"TOP_TITLES_ER_{idx}_VALUE"] = er_row["value"]

        mapped[f"TOP_TITLES_VCR_{idx}_NAME"] = vcr_row["name"]
        mapped[f"TOP_TITLES_VCR_{idx}_VALUE"] = vcr_row["value"]

    return mapped


def parse_eoc(file_path: Path) -> Dict[str, Any]:
    mapped = build_mapped_values(file_path)

    missing = [
        key for key in [
            "CAMPAIGN_NAME",
            "IMPRESSIONS",
            "CLICKS",
            "CTR",
        ]
        if not mapped.get(key)
    ]

    return {
        "file_name": file_path.name,
        "version": APP_VERSION,
        "template_source": str(TEMPLATE_PATH),
        "status": "ready" if not missing else "needs_review",
        "missing_critical_fields": missing,
        "mapped_values": mapped,
    }


# -----------------------------
# PowerPoint generation
# -----------------------------

def replace_in_text_frame(text_frame, values: Dict[str, str]) -> None:
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            text = run.text
            for key, value in values.items():
                text = text.replace(f"{{{{{key}}}}}", clean_text(value))
            run.text = text


def replace_placeholders_in_shape(shape, values: Dict[str, str]) -> None:
    if hasattr(shape, "text_frame") and shape.has_text_frame:
        replace_in_text_frame(shape.text_frame, values)

    if hasattr(shape, "table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_in_text_frame(cell.text_frame, values)

    if hasattr(shape, "shapes"):
        for child in shape.shapes:
            replace_placeholders_in_shape(child, values)


def generate_ppt(mapped_values: Dict[str, str], output_path: Path) -> None:
    if not TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=f"Template not found at {TEMPLATE_PATH}. Confirm templates/exec_summary_master.pptx exists in repo.",
        )

    prs = Presentation(str(TEMPLATE_PATH))

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_placeholders_in_shape(shape, mapped_values)

    prs.save(str(output_path))


def create_output_filename(mapped_values: Dict[str, str]) -> str:
    client = mapped_values.get("CLIENT_NAME") or "Client"
    campaign = mapped_values.get("CAMPAIGN_NAME") or "Campaign"
    raw = f"{client} - {campaign} - Exec Summary.pptx"
    return re.sub(r"[^A-Za-z0-9._ -]+", "_", raw).strip()


# -----------------------------
# API endpoints
# -----------------------------

@app.get("/health")
def health():
    return {
        "status": "ok",
        "version": APP_VERSION,
        "template_exists": TEMPLATE_PATH.exists(),
        "template_path": str(TEMPLATE_PATH),
    }


@app.post("/validate-eoc")
async def validate_eoc(payload: FileRefsPayload):
    if not payload.openaiFileIdRefs:
        raise HTTPException(status_code=400, detail="No EOC file supplied.")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        eoc_path = await download_openai_file(payload.openaiFileIdRefs[0], tmp_dir)
        validation = parse_eoc(eoc_path)

    return JSONResponse(validation)


@app.post("/generate-exec-summary")
async def generate_exec_summary(payload: FileRefsPayload):
    if not payload.openaiFileIdRefs:
        raise HTTPException(status_code=400, detail="No EOC file supplied.")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        eoc_path = await download_openai_file(payload.openaiFileIdRefs[0], tmp_dir)

        validation = parse_eoc(eoc_path)
        mapped_values = validation["mapped_values"]

        filename = create_output_filename(mapped_values)
        pptx_path = tmp_dir / filename

        generate_ppt(mapped_values, pptx_path)

        file_size = pptx_path.stat().st_size
        if file_size > MAX_RETURN_FILE_BYTES:
            raise HTTPException(
                status_code=413,
                detail="Generated PPT is over 10MB. Reduce template media size.",
            )

        encoded = base64.b64encode(pptx_path.read_bytes()).decode("utf-8")

    return {
        "status": "success",
        "version": APP_VERSION,
        "summary": validation,
        "openaiFileResponse": [
            {
                "name": filename,
                "mime_type": PPTX_MIME,
                "content": encoded,
            }
        ],
    }
