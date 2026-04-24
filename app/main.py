import base64
import json
import os
import re
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import httpx
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from pptx import Presentation


APP_VERSION = "10.0.4-clean-production"

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = Path(os.getenv("TEMPLATE_PATH", BASE_DIR / "templates" / "exec_summary_master.pptx"))

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MAX_RETURN_FILE_BYTES = 10 * 1024 * 1024


app = FastAPI(
    title="PCA Automation API",
    version=APP_VERSION,
    description="Validates EOC Excel files and generates editable Exec Summary PPT decks.",
)


class OpenAIFileRef(BaseModel):
    id: Optional[str] = None
    name: Optional[str] = None
    mime_type: Optional[str] = None
    download_link: Optional[str] = None


class FileRefsPayload(BaseModel):
    openaiFileIdRefs: List[OpenAIFileRef] = Field(
        ...,
        description="Array of files supplied by OpenAI Actions. Upload exactly one EOC Excel file.",
    )


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalise_key(value: str) -> str:
    value = clean_text(value).lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    return value.strip("_")


def fmt_int(value: Any) -> str:
    try:
        return f"{int(round(float(value))):,}"
    except Exception:
        return clean_text(value)


def fmt_percent(value: Any) -> str:
    if value is None or clean_text(value) == "":
        return ""
    try:
        num = float(str(value).replace("%", "").strip())
        if 0 < num <= 1:
            num *= 100
        return f"{num:.2f}%"
    except Exception:
        return clean_text(value)


def parse_date(value: Any) -> Optional[datetime]:
    if value is None or clean_text(value) == "":
        return None
    try:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.to_pydatetime()
    except Exception:
        return None


def format_date_range_short(start: Optional[datetime], end: Optional[datetime]) -> str:
    if not start and not end:
        return ""
    if start and not end:
        return start.strftime("%d %b").lstrip("0")
    if end and not start:
        return end.strftime("%d %b").lstrip("0")

    if start.year == end.year:
        if start.month == end.month:
            return f"{start.day} - {end.day} {end.strftime('%b')}"
        return f"{start.day} {start.strftime('%b')} - {end.day} {end.strftime('%b')}"
    return f"{start.day} {start.strftime('%b %Y')} - {end.day} {end.strftime('%b %Y')}"


def format_date_range_full(start: Optional[datetime], end: Optional[datetime]) -> str:
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


async def download_openai_file(file_ref: OpenAIFileRef, dest_dir: Path) -> Path:
    if not file_ref.download_link:
        raise HTTPException(status_code=400, detail="Missing download_link in openaiFileIdRefs.")

    safe_name = file_ref.name or "uploaded_eoc.xlsx"
    safe_name = re.sub(r"[^A-Za-z0-9._ -]+", "_", safe_name)
    output_path = dest_dir / safe_name

    async with httpx.AsyncClient(timeout=30.0, follow_redirects=True) as client:
        response = await client.get(file_ref.download_link)
        response.raise_for_status()
        output_path.write_bytes(response.content)

    return output_path


def read_excel_sheets(file_path: Path) -> Dict[str, pd.DataFrame]:
    try:
        return pd.read_excel(file_path, sheet_name=None, header=None, engine="openpyxl")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not read EOC Excel file: {exc}")


def build_flat_cell_map(sheets: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    found: Dict[str, Any] = {}

    for _, df in sheets.items():
        rows, cols = df.shape

        for r in range(rows):
            for c in range(cols):
                raw_key = clean_text(df.iat[r, c])
                if not raw_key:
                    continue

                key = normalise_key(raw_key)

                right_val = df.iat[r, c + 1] if c + 1 < cols else None
                below_val = df.iat[r + 1, c] if r + 1 < rows else None

                if clean_text(right_val):
                    found[key] = right_val
                elif clean_text(below_val):
                    found[key] = below_val

    return found


def find_first(cell_map: Dict[str, Any], candidates: List[str]) -> Any:
    for candidate in candidates:
        norm = normalise_key(candidate)
        if norm in cell_map and clean_text(cell_map[norm]):
            return cell_map[norm]

    for key, value in cell_map.items():
        for candidate in candidates:
            if normalise_key(candidate) in key and clean_text(value):
                return value

    return ""


def extract_table_rows(sheets: Dict[str, pd.DataFrame]) -> List[Dict[str, Any]]:
    all_rows: List[Dict[str, Any]] = []

    for _, df in sheets.items():
        for header_row_idx in range(min(len(df), 80)):
            headers = [normalise_key(x) for x in df.iloc[header_row_idx].tolist()]
            if not any(h in headers for h in ["site", "publisher", "title", "placement"]):
                continue

            for r in range(header_row_idx + 1, len(df)):
                row_data = {}
                blank_count = 0

                for c, header in enumerate(headers):
                    if not header:
                        continue
                    value = df.iat[r, c] if c < df.shape[1] else None
                    if clean_text(value):
                        row_data[header] = value
                    else:
                        blank_count += 1

                if row_data:
                    all_rows.append(row_data)

    return all_rows


def top_rows_by_metric(rows: List[Dict[str, Any]], metric_candidates: List[str], limit: int = 3) -> List[Dict[str, str]]:
    metric_keys = [normalise_key(x) for x in metric_candidates]
    name_keys = ["site", "publisher", "title", "placement", "domain"]

    scored = []

    for row in rows:
        name = ""
        for key in name_keys:
            if key in row and clean_text(row[key]):
                name = clean_text(row[key])
                break

        if not name:
            continue

        metric_value = None
        for key in metric_keys:
            if key in row and clean_text(row[key]):
                metric_value = row[key]
                break

        if metric_value is None:
            continue

        try:
            numeric = float(str(metric_value).replace("%", "").replace(",", "").strip())
        except Exception:
            continue

        scored.append((numeric, name, metric_value))

    scored.sort(reverse=True, key=lambda x: x[0])

    return [
        {"name": name, "value": fmt_percent(value) if "rate" in "_".join(metric_keys) or "ctr" in "_".join(metric_keys) else fmt_int(value)}
        for _, name, value in scored[:limit]
    ]


def parse_eoc(file_path: Path) -> Dict[str, Any]:
    sheets = read_excel_sheets(file_path)
    cell_map = build_flat_cell_map(sheets)
    rows = extract_table_rows(sheets)

    start_date = parse_date(find_first(cell_map, ["Start Date", "Live Start", "Campaign Start Date", "From"]))
    end_date = parse_date(find_first(cell_map, ["End Date", "Live End", "Campaign End Date", "To"]))

    campaign_name = find_first(cell_map, ["Campaign Name", "Campaign", "Name"])
    client_name = find_first(cell_map, ["Client", "Advertiser", "Brand"])
    market = find_first(cell_map, ["Market", "Markets", "Country", "Countries"])
    format_name = find_first(cell_map, ["Format", "Product", "Creative Format"])

    impressions = find_first(cell_map, ["Impressions", "Delivered Impressions", "Total Impressions"])
    clicks = find_first(cell_map, ["Clicks", "Total Clicks"])
    engagements = find_first(cell_map, ["Engagements", "Total Engagements"])
    completed_views = find_first(cell_map, ["Completed Views", "Video Completes", "Completions"])

    ctr = find_first(cell_map, ["CTR", "Click Through Rate", "Click-Through Rate"])
    engagement_rate = find_first(cell_map, ["Engagement Rate", "ER"])
    viewability = find_first(cell_map, ["Viewability", "Viewability Rate"])
    vcr = find_first(cell_map, ["VCR", "Video Completion Rate", "Completion Rate"])

    top_ctr = top_rows_by_metric(rows, ["CTR", "Click Through Rate"], 3)
    top_er = top_rows_by_metric(rows, ["Engagement Rate", "ER"], 3)

    mapped = {
        "CAMPAIGN_NAME": clean_text(campaign_name),
        "CLIENT_NAME": clean_text(client_name),
        "MARKETS": clean_text(market),
        "FORMAT": clean_text(format_name),
        "LIVE_DATES_SHORT": format_date_range_short(start_date, end_date),
        "LIVE_DATES_FULL": format_date_range_full(start_date, end_date),
        "IMPRESSIONS": fmt_int(impressions),
        "CLICKS": fmt_int(clicks),
        "ENGAGEMENTS": fmt_int(engagements),
        "COMPLETED_VIEWS": fmt_int(completed_views),
        "CTR": fmt_percent(ctr),
        "ENGAGEMENT_RATE": fmt_percent(engagement_rate),
        "VIEWABILITY": fmt_percent(viewability),
        "VCR": fmt_percent(vcr),
        "VIDEO_COMPLETION_RATE": fmt_percent(vcr),
    }

    for idx in range(1, 4):
        ctr_row = top_ctr[idx - 1] if len(top_ctr) >= idx else {"name": "", "value": ""}
        er_row = top_er[idx - 1] if len(top_er) >= idx else {"name": "", "value": ""}

        mapped[f"TOP_TITLES_CTR_{idx}_NAME"] = ctr_row["name"]
        mapped[f"TOP_TITLES_CTR_{idx}_VALUE"] = ctr_row["value"]
        mapped[f"TOP_TITLES_ER_{idx}_NAME"] = er_row["name"]
        mapped[f"TOP_TITLES_ER_{idx}_VALUE"] = er_row["value"]

    validation = {
        "file_name": file_path.name,
        "version": APP_VERSION,
        "template_source": str(TEMPLATE_PATH),
        "mapped_values": mapped,
        "missing_critical_fields": [
            key for key in ["CAMPAIGN_NAME", "CLIENT_NAME", "MARKETS", "CTR"]
            if not mapped.get(key)
        ],
        "status": "ready" if mapped.get("CAMPAIGN_NAME") or mapped.get("CLIENT_NAME") else "needs_review",
    }

    return validation


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
            detail=f"Template not found at {TEMPLATE_PATH}. Confirm /templates/exec_summary_master.pptx exists in repo.",
        )

    prs = Presentation(str(TEMPLATE_PATH))

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_placeholders_in_shape(shape, mapped_values)

    prs.save(str(output_path))


def create_output_filename(mapped_values: Dict[str, str]) -> str:
    client = mapped_values.get("CLIENT_NAME") or "Client"
    campaign = mapped_values.get("CAMPAIGN_NAME") or "Exec Summary"
    raw = f"{client} - {campaign} - Exec Summary.pptx"
    return re.sub(r"[^A-Za-z0-9._ -]+", "_", raw).strip()


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

        size = pptx_path.stat().st_size
        if size > MAX_RETURN_FILE_BYTES:
            raise HTTPException(
                status_code=413,
                detail="Generated PPT is over 10MB. Reduce template media size or switch to URL return mode.",
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
