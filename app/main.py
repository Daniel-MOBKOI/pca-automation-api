import base64
import math
import os
import re
import tempfile
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import httpx
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from pptx import Presentation


APP_VERSION = "10.1.2-hybrid-polish-na"
MISSING_VALUE = "N/A (not specified in source file)"

app = FastAPI(title="PCA Automation API", version=APP_VERSION)

BASE_DIR = Path(__file__).resolve().parent
ROOT_DIR = BASE_DIR.parent

TEMPLATE_PATH = Path(os.getenv("TEMPLATE_PATH", ROOT_DIR / "templates" / "exec_summary_master.pptx"))
RULES_MASTER_PATH = Path(os.getenv("RULES_MASTER_PATH", ROOT_DIR / "assets" / "PCA_GPT_Rules_Master.xlsx"))

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MAX_RETURN_FILE_BYTES = 10 * 1024 * 1024


class OpenAIFileRef(BaseModel):
    id: Optional[str] = None
    name: Optional[str] = None
    mime_type: Optional[str] = None
    download_link: Optional[str] = None


class FileRefsPayload(BaseModel):
    openaiFileIdRefs: List[OpenAIFileRef] = Field(...)


@app.get("/health")
def health():
    return {
        "status": "ok",
        "version": APP_VERSION,
        "template_exists": TEMPLATE_PATH.exists(),
        "template_path": str(TEMPLATE_PATH),
        "rules_master_exists": RULES_MASTER_PATH.exists(),
        "rules_master_path": str(RULES_MASTER_PATH),
    }


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value)).strip()


def clean_dimension_label(value: Any) -> str:
    value = clean_text(value)
    value = re.sub(r"\bMISCRL\b", "", value, flags=re.IGNORECASE)
    value = re.sub(r"([a-z])([A-Z])", r"\1 \2", value)
    value = re.sub(r"\bDisplay\s*Animated\b", "Display Animated", value, flags=re.IGNORECASE)
    value = re.sub(r"\bIn\s*Situ\s*Video\b", "In Situ Video", value, flags=re.IGNORECASE)
    value = re.sub(r"[-_]+", " - ", value)
    value = re.sub(r"\s+-\s+-\s+", " - ", value)
    value = re.sub(r"\s{2,}", " ", value)
    return value.strip(" -")


def fill_missing_mapped_values(mapped: Dict[str, Any]) -> Dict[str, Any]:
    for key, value in mapped.items():
        if value is None or clean_text(value) == "":
            mapped[key] = MISSING_VALUE
    return mapped


MARKET_CODE_MAP = {
    "france": "FR", "french": "FR", "fr": "FR",
    "united kingdom": "UK", "uk": "UK", "gb": "UK", "great britain": "UK", "britain": "UK",
    "italy": "IT", "italian": "IT", "it": "IT",
    "spain": "ES", "spanish": "ES", "es": "ES",
    "germany": "DE", "german": "DE", "de": "DE",
    "netherlands": "NL", "dutch": "NL", "nl": "NL",
    "belgium": "BE", "be": "BE",
    "switzerland": "CH", "ch": "CH",
    "austria": "AT", "at": "AT",
    "portugal": "PT", "pt": "PT",
    "ireland": "IE", "ie": "IE",
    "usa": "US", "us": "US", "united states": "US", "united states of america": "US",
    "canada": "CA", "ca": "CA",
    "australia": "AU", "au": "AU",
    "japan": "JP", "jp": "JP",
    "hong kong": "HK", "hk": "HK",
    "singapore": "SG", "sg": "SG",
    "thailand": "TH", "thai": "TH", "th": "TH",
    "taiwan": "TW", "tw": "TW",
    "china": "CN", "cn": "CN",
    "korea": "KR", "south korea": "KR", "kr": "KR",
    "india": "IN", "in": "IN",
    "mexico": "MX", "mx": "MX",
    "brazil": "BR", "br": "BR",
}


def normalise_market_label(value: Any) -> str:
    raw = clean_dimension_label(value)
    if not raw:
        return ""

    raw = re.sub(r"\b(total|totals|overall|summary|market|markets|country|countries)\b", "", raw, flags=re.IGNORECASE)
    raw = raw.strip(" ,-")

    if not raw:
        return ""

    parts = re.split(r"[,;/|]+", raw)
    out = []

    for part in parts:
        part = clean_text(part).strip(" -")
        if not part:
            continue

        key = part.lower()
        code = MARKET_CODE_MAP.get(key)

        if not code:
            if re.fullmatch(r"[A-Za-z]{2,3}", part):
                code = part.upper()
            else:
                code = part.upper() if len(part) <= 3 else part

        if code and code not in out:
            out.append(code)

    return ", ".join(out)


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


def abs_number(value: Any) -> Optional[float]:
    num = safe_number(value)
    if num is None:
        return None
    return abs(num)


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
            py_dt = dt.to_pydatetime()
            if 2000 <= py_dt.year <= 2100:
                return py_dt
    except Exception:
        pass

    match = re.search(r"\b(20\d{2}-\d{2}-\d{2})\b", text)
    if match:
        try:
            return datetime.strptime(match.group(1), "%Y-%m-%d")
        except Exception:
            pass

    match = re.search(r"\b(\d{1,2}[/-]\d{1,2}[/-]20\d{2})\b", text)
    if match:
        try:
            dt = pd.to_datetime(match.group(1), dayfirst=True, errors="coerce")
            if pd.notna(dt):
                return dt.to_pydatetime()
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


def is_blank_row(values: List[Any]) -> bool:
    return all(clean_text(v) == "" for v in values)


def load_rules_master() -> Optional[Dict[str, pd.DataFrame]]:
    if not RULES_MASTER_PATH.exists():
        return None

    try:
        xls = pd.ExcelFile(RULES_MASTER_PATH, engine="openpyxl")
        out = {}
        for sheet in xls.sheet_names:
            out[sheet] = pd.read_excel(RULES_MASTER_PATH, sheet_name=sheet, engine="openpyxl")
        return out
    except Exception:
        return None


ALIASES = {
    "campaign": ["campaign", "campaign name"],
    "client_brand": ["client", "brand", "advertiser", "brand / client"],
    "impressions": ["impressions", "delivered impressions", "served impressions"],
    "ctr": ["ctr", "click through rate", "click-through rate"],
    "engagement_rate": ["engagement rate", "engagement %", "er", "total er", "overall er"],
    "vcr": ["video completion rate", "vcr", "completed view rate", "video % complete"],
    "on_screen": ["mobkoi on screen", "on screen", "on-screen", "mrc viewability", "viewability", "viewable rate"],
    "mobkoi_on_screen": ["mobkoi on screen"],
    "mrc_viewability": ["mrc viewability"],
    "spend": ["actual spend", "spend", "total spend", "media spend"],
    "sold_paid_units": ["sold paid units"],
    "delivered_overall_av_units": ["delivered overall av units", "delivered overall av (units)"],
    "delivery_incl_av": ["delivery percentage incl av", "delivery percentage (incl av)"],
    "delivered_av_amount": ["delivered av amount currency", "delivered av amount (currency)", "worth of added value"],
    "site": ["site", "publisher", "domain", "environment", "property", "placement", "title", "inventory", "app", "website"],
    "geo": ["geo", "market", "markets", "country", "countries", "region", "territory", "location", "locale"],
    "format": ["format", "formats", "creative", "creative format", "ad format", "unit type", "product", "product type", "placement type"],
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


def score_header_row(row_values: List[Any]) -> int:
    keywords = [
        "campaign", "impressions", "ctr", "engagement", "vcr", "completion",
        "spend", "site", "publisher", "domain", "environment", "property",
        "placement", "title", "inventory", "app", "website",
        "geo", "market", "markets", "country", "countries", "region", "territory", "location", "locale",
        "format", "formats", "creative", "product", "unit type",
        "date", "report date", "served date", "live date",
        "sold paid units", "delivered overall av", "delivery percentage", "mrc viewability",
    ]

    score = 0
    for cell in row_values:
        text = normalize_text(cell)
        for kw in keywords:
            if kw in text:
                score += 1

    return score


def find_candidate_header_rows(df: pd.DataFrame) -> List[int]:
    candidates = []

    for i in range(min(120, len(df))):
        row_values = df.iloc[i].tolist()
        score = score_header_row(row_values)

        if score >= 3:
            candidates.append(i)

    deduped = []
    for idx in candidates:
        if not deduped or idx - deduped[-1] > 2:
            deduped.append(idx)

    return deduped


def build_block_from_header(df: pd.DataFrame, header_row: int) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    temp = df.iloc[header_row:].copy()
    if temp.empty:
        return pd.DataFrame()

    temp.columns = [clean_text(c) for c in temp.iloc[0]]
    temp = temp[1:].reset_index(drop=True)
    temp = temp.dropna(axis=1, how="all")

    rows = []
    blank_streak = 0

    for _, row in temp.iterrows():
        row_vals = row.tolist()

        if is_blank_row(row_vals):
            blank_streak += 1
            if blank_streak >= 1:
                break
            continue

        blank_streak = 0
        rows.append(row_vals)

    if not rows:
        return pd.DataFrame(columns=temp.columns)

    block = pd.DataFrame(rows, columns=temp.columns)

    if not block.empty:
        header_norm = [normalize_header(c) for c in block.columns]
        keep_rows = []

        for _, row in block.iterrows():
            row_norm = [normalize_header(v) for v in row.tolist()]
            same_count = 0

            for a, b in zip(header_norm[:12], row_norm[:12]):
                if a and b and a == b:
                    same_count += 1

            keep_rows.append(same_count < 3)

        block = block[pd.Series(keep_rows).values].reset_index(drop=True)

    return block


def classify_table(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "unknown"

    site_col = find_col(df, "site")
    geo_col = find_col(df, "geo")
    format_col = find_col(df, "format")

    has_ctr = find_col(df, "ctr") is not None
    has_er = find_col(df, "engagement_rate") is not None
    has_vcr = find_col(df, "vcr") is not None
    has_impressions = find_col(df, "impressions") is not None

    delivery_hits = sum([
        1 if find_col(df, "sold_paid_units") else 0,
        1 if find_col(df, "delivered_overall_av_units") else 0,
        1 if find_col(df, "delivery_incl_av") else 0,
        1 if find_col(df, "delivered_av_amount") else 0,
    ])

    if delivery_hits >= 2:
        return "campaign_delivery"

    if site_col and (has_ctr or has_er or has_vcr):
        return "site"

    if geo_col and has_impressions:
        return "geo"

    if format_col and has_impressions:
        return "format"

    date_col = find_col(df, "date")
    if date_col:
        col_data = df[date_col]
        if isinstance(col_data, pd.DataFrame):
            col_data = col_data.iloc[:, 0]

        sample = [parse_date_from_any(v) for v in col_data.head(10).values]
        if any(v is not None for v in sample):
            return "date"

    kpi_hits = sum([
        1 if find_col(df, "campaign") else 0,
        1 if has_impressions else 0,
        1 if has_ctr else 0,
        1 if has_er else 0,
        1 if has_vcr else 0,
        1 if (find_col(df, "mobkoi_on_screen") or find_col(df, "mrc_viewability") or find_col(df, "on_screen")) else 0,
        1 if find_col(df, "spend") else 0,
    ])

    if kpi_hits >= 3:
        return "campaign_kpi_summary"

    if site_col:
        return "site"

    if geo_col:
        return "geo"

    if format_col:
        return "format"

    return "unknown"


def table_quality_score(df: pd.DataFrame, table_type: str) -> int:
    if df is None or df.empty:
        return -999

    score = min(len(df.columns), 20) + min(len(df), 10)

    if table_type == "campaign_kpi_summary":
        for key in ["campaign", "impressions", "ctr", "engagement_rate", "vcr", "spend"]:
            if find_col(df, key):
                score += 10
        if find_col(df, "mobkoi_on_screen"):
            score += 12
        elif find_col(df, "mrc_viewability"):
            score += 6

    elif table_type == "campaign_delivery":
        for key in ["sold_paid_units", "delivered_overall_av_units", "delivery_incl_av", "delivered_av_amount"]:
            if find_col(df, key):
                score += 12

    elif table_type == "site":
        for key in ["site", "ctr", "engagement_rate", "vcr"]:
            if find_col(df, key):
                score += 10

    elif table_type == "geo":
        for key in ["geo", "impressions"]:
            if find_col(df, key):
                score += 10

    elif table_type == "format":
        for key in ["format", "impressions"]:
            if find_col(df, key):
                score += 10

    elif table_type == "date":
        if find_col(df, "date"):
            score += 20

    return score


def detect_tables(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    candidates_by_type: Dict[str, List[Tuple[int, pd.DataFrame]]] = {
        "campaign_kpi_summary": [],
        "campaign_delivery": [],
        "site": [],
        "geo": [],
        "format": [],
        "date": [],
    }

    for _, raw_df in sheets.items():
        if raw_df is None or raw_df.empty:
            continue

        for header_row in find_candidate_header_rows(raw_df):
            block = build_block_from_header(raw_df, header_row)
            if block.empty:
                continue

            table_type = classify_table(block)
            if table_type == "unknown":
                continue

            score = table_quality_score(block, table_type)
            candidates_by_type[table_type].append((score, block))

    detected: Dict[str, pd.DataFrame] = {}
    for table_type, candidates in candidates_by_type.items():
        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            detected[table_type] = candidates[0][1]

    return detected


def scan_workbook_texts(sheets: Dict[str, pd.DataFrame]) -> List[str]:
    texts = []

    for df in sheets.values():
        if df is None or df.empty:
            continue

        for row in df.head(20).values:
            for cell in row:
                text = clean_text(cell)
                if text:
                    texts.append(text)

    return texts


def extract_dates(
    sheets: Dict[str, pd.DataFrame],
    detected: Dict[str, pd.DataFrame],
    filename: str = "",
) -> Tuple[Optional[datetime], Optional[datetime]]:
    date_df = detected.get("date")

    if date_df is not None:
        date_col = find_col(date_df, "date")
        if date_col:
            col_data = date_df[date_col]
            if isinstance(col_data, pd.DataFrame):
                col_data = col_data.iloc[:, 0]

            vals = [parse_date_from_any(v) for v in col_data.tolist()]
            vals = [v for v in vals if v]
            if vals:
                return min(vals), max(vals)

    for text in scan_workbook_texts(sheets):
        low = normalize_text(text)
        if "report date" in low:
            dt = parse_date_from_any(text)
            if dt:
                return dt - timedelta(days=6), dt

    match = re.search(r"\b(20\d{2}-\d{2}-\d{2})\b", filename or "")
    if match:
        end_dt = parse_date_from_any(match.group(1))
        if end_dt:
            return end_dt - timedelta(days=6), end_dt

    return None, None


def extract_kpi_value(df: pd.DataFrame, key: str) -> Any:
    col = find_col(df, key)
    if not col:
        return None

    col_data = df[col]
    if isinstance(col_data, pd.DataFrame):
        col_data = col_data.iloc[:, 0]

    return first_non_empty(col_data)


def extract_client_name(
    sheets: Dict[str, pd.DataFrame],
    detected: Dict[str, pd.DataFrame],
    filename: str = "",
) -> Optional[str]:
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

    kpi_df = detected.get("campaign_kpi_summary")
    if kpi_df is not None:
        campaign_val = clean_text(extract_kpi_value(kpi_df, "campaign"))
        if campaign_val:
            for sep in [" - ", " | ", "_"]:
                if sep in campaign_val:
                    return clean_text(campaign_val.split(sep)[0])
            words = campaign_val.split()
            if words:
                return words[0]

    if filename:
        base = Path(filename).stem
        if " - " in base:
            return clean_text(base.split(" - ")[0])

    return None


def join_dimension_values(df: pd.DataFrame, dim_key: str) -> Optional[str]:
    col = find_col(df, dim_key)
    if not col:
        return None

    col_data = df[col]
    if isinstance(col_data, pd.DataFrame):
        col_data = col_data.iloc[:, 0]

    if dim_key == "geo":
        raw_vals = []
        for v in col_data.tolist():
            cleaned = normalise_market_label(v)
            if cleaned:
                raw_vals.extend([x.strip() for x in cleaned.split(",") if x.strip()])
    else:
        raw_vals = [clean_dimension_label(v) for v in col_data.tolist() if clean_dimension_label(v)]

    blocked = {"total", "totals", "overall", "summary", "campaign", "format", "formats", "market", "markets"}
    raw_vals = [v for v in raw_vals if normalize_header(v) not in blocked]

    if not raw_vals:
        return None

    base = raw_vals[0]
    cleaned = [base]

    for val in raw_vals[1:]:
        if dim_key == "geo":
            candidate = val
        else:
            parts_base = base.split(" - ")
            parts_val = val.split(" - ")

            suffix = parts_val

            for i in range(min(len(parts_base), len(parts_val))):
                if parts_base[i] != parts_val[i]:
                    suffix = parts_val[i:]
                    break
            else:
                if len(parts_val) > len(parts_base):
                    suffix = parts_val[len(parts_base):]
                else:
                    suffix = [val]

            candidate = clean_dimension_label(" - ".join([p for p in suffix if p]))

        if candidate:
            cleaned.append(candidate)

    seen = []
    final = []

    for v in cleaned:
        if v not in seen:
            seen.append(v)
            final.append(v)

    return ", ".join(final)


def extract_top_titles(df: pd.DataFrame, metric_key: str, max_rank: int = 5) -> Dict[str, Any]:
    site_col = find_col(df, "site")
    metric_col = find_col(df, metric_key)

    if not site_col or not metric_col:
        return {}

    site_data = df[site_col]
    if isinstance(site_data, pd.DataFrame):
        site_data = site_data.iloc[:, 0]

    metric_data = df[metric_col]
    if isinstance(metric_data, pd.DataFrame):
        metric_data = metric_data.iloc[:, 0]

    temp = pd.DataFrame({"site": site_data, "metric": metric_data})
    temp["site_clean"] = temp["site"].apply(clean_text)
    temp["site_norm"] = temp["site_clean"].apply(lambda x: normalize_header(x))
    temp["metric"] = temp["metric"].apply(safe_number)
    temp = temp.dropna(subset=["metric"])

    blocked_exact = {
        "site", "publisher", "domain", "environment", "property",
        "placement", "title", "inventory", "app", "website",
        "total", "totals", "overall", "average", "avg", "summary",
        "campaign", "campaign name", "market", "format", "creative"
    }

    temp = temp[~temp["site_norm"].isin(blocked_exact)]
    temp = temp[temp["site_clean"] != ""]

    if temp.empty:
        return {}

    temp = temp.sort_values(by="metric", ascending=False).head(max_rank)

    prefix = {
        "ctr": "TOP_TITLES_CTR",
        "engagement_rate": "TOP_TITLES_ER",
        "vcr": "TOP_TITLES_VCR",
    }[metric_key]

    out = {}

    for idx, (_, row) in enumerate(temp.iterrows(), start=1):
        out[f"{prefix}_{idx}_NAME"] = row["site_clean"]
        out[f"{prefix}_{idx}_VALUE"] = format_percent(row["metric"])

    return out


def validate_mapped_values(mapped: Dict[str, Any]) -> Dict[str, Any]:
    required_core = [
        "CAMPAIGN_NAME",
        "DELIVERED_IMPRESSIONS",
        "PERFORMANCE_CTR",
        "PERFORMANCE_ENGAGEMENT_RATE",
        "PERFORMANCE_VCR",
        "LIVE_DATES_FULL",
    ]

    missing_required = [
        k for k in required_core
        if not mapped.get(k) or mapped.get(k) == MISSING_VALUE
    ]

    warnings = []
    if mapped.get("CLIENT_NAME") == MISSING_VALUE:
        warnings.append("CLIENT_NAME missing")
    if mapped.get("CAMPAIGN_MARKETS") == MISSING_VALUE:
        warnings.append("CAMPAIGN_MARKETS missing")
    if mapped.get("CAMPAIGN_FORMATS") == MISSING_VALUE:
        warnings.append("CAMPAIGN_FORMATS missing")
    if mapped.get("TOP_TITLES_CTR_1_NAME") == MISSING_VALUE:
        warnings.append("Top Titles CTR missing")
    if mapped.get("TOP_TITLES_VCR_1_NAME") == MISSING_VALUE:
        warnings.append("Top Titles VCR missing")
    if mapped.get("TOP_TITLES_ER_1_NAME") == MISSING_VALUE:
        warnings.append("Top Titles ER missing")

    return {
        "is_valid": len(missing_required) == 0,
        "missing_required": missing_required,
        "warnings": warnings,
    }


def build_mapped_values(sheets: Dict[str, pd.DataFrame], filename: str = "") -> Dict[str, Any]:
    rules = load_rules_master()
    detected = detect_tables(sheets)

    kpi_df = detected.get("campaign_kpi_summary")
    delivery_df = detected.get("campaign_delivery")
    site_df = detected.get("site")
    geo_df = detected.get("geo")
    format_df = detected.get("format")

    start_dt, end_dt = extract_dates(sheets, detected, filename)
    client_name = extract_client_name(sheets, detected, filename)

    mapped: Dict[str, Any] = {}

    if kpi_df is not None:
        mapped["CAMPAIGN_NAME"] = clean_text(extract_kpi_value(kpi_df, "campaign"))
        mapped["DELIVERED_IMPRESSIONS"] = format_number(safe_number(extract_kpi_value(kpi_df, "impressions")), 0)
        mapped["PERFORMANCE_CTR"] = format_percent(extract_kpi_value(kpi_df, "ctr"))
        mapped["PERFORMANCE_ENGAGEMENT_RATE"] = format_percent(extract_kpi_value(kpi_df, "engagement_rate"))
        mapped["PERFORMANCE_VCR"] = format_percent(extract_kpi_value(kpi_df, "vcr"))

        on_screen_val = None
        for key in ["mobkoi_on_screen", "mrc_viewability", "on_screen"]:
            col = find_col(kpi_df, key)
            if col:
                col_data = kpi_df[col]
                if isinstance(col_data, pd.DataFrame):
                    col_data = col_data.iloc[:, 0]
                on_screen_val = first_non_empty(col_data)
                break

        mapped["PERFORMANCE_ON_SCREEN"] = format_percent(on_screen_val)
        mapped["CAMPAIGN_BUDGET"] = format_currency(safe_number(extract_kpi_value(kpi_df, "spend")))
    else:
        mapped["CAMPAIGN_NAME"] = None
        mapped["DELIVERED_IMPRESSIONS"] = None
        mapped["PERFORMANCE_CTR"] = None
        mapped["PERFORMANCE_ENGAGEMENT_RATE"] = None
        mapped["PERFORMANCE_VCR"] = None
        mapped["PERFORMANCE_ON_SCREEN"] = None
        mapped["CAMPAIGN_BUDGET"] = None

    if delivery_df is not None:
        io_val = safe_number(extract_kpi_value(delivery_df, "sold_paid_units"))
        av_units_val = abs_number(extract_kpi_value(delivery_df, "delivered_overall_av_units"))
        av_amount_val = abs_number(extract_kpi_value(delivery_df, "delivered_av_amount"))

        mapped["IO_OVERALL_IMPRESSIONS"] = format_number(io_val, 0)
        mapped["DELIVERED_OVERALL_AV_UNITS"] = format_number(av_units_val, 0)
        mapped["DELIVERY_WITH_AV_PERCENT"] = format_percent(extract_kpi_value(delivery_df, "delivery_incl_av"))
        mapped["ADDED_VALUE_WORTH"] = format_currency(av_amount_val)
        mapped["ADDED_VALUE_IMPRESSIONS"] = format_number(av_units_val, 0)
    else:
        mapped["IO_OVERALL_IMPRESSIONS"] = None
        mapped["DELIVERED_OVERALL_AV_UNITS"] = None
        mapped["DELIVERY_WITH_AV_PERCENT"] = None
        mapped["ADDED_VALUE_WORTH"] = None
        mapped["ADDED_VALUE_IMPRESSIONS"] = None

    mapped["CLIENT_NAME"] = client_name
    mapped["CLIENT"] = client_name
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

    mapped = fill_missing_mapped_values(mapped)
    validation = validate_mapped_values(mapped)

    diagnostics = {
        "detected_tables": list(detected.keys()),
        "table_columns": {k: [str(c) for c in v.columns] for k, v in detected.items()},
        "rules_master_loaded": rules is not None,
        "version": APP_VERSION,
    }

    return {
        "mapped_values": mapped,
        "validation": validation,
        "diagnostics": diagnostics,
    }


async def download_openai_file(file_ref: OpenAIFileRef, dest_dir: Path) -> Path:
    if not file_ref.download_link:
        raise HTTPException(status_code=400, detail="Missing download_link in openaiFileIdRefs.")

    safe_name = file_ref.name or "uploaded_eoc.xlsx"
    safe_name = re.sub(r"[^A-Za-z0-9._ -]+", "_", safe_name)
    output_path = dest_dir / safe_name

    async with httpx.AsyncClient(timeout=60.0, follow_redirects=True) as client:
        response = await client.get(file_ref.download_link)
        response.raise_for_status()
        output_path.write_bytes(response.content)

    return output_path


def read_uploaded_eoc(path: Path) -> Dict[str, pd.DataFrame]:
    try:
        return pd.read_excel(path, sheet_name=None, header=None, engine="openpyxl")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not read EOC Excel file: {exc}")


def replace_text(text: str, data: dict) -> str:
    if not text:
        return text

    for key, value in data.items():
        text = text.replace(f"{{{{{key}}}}}", str(value))

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

    if hasattr(shape, "shapes"):
        for child in shape.shapes:
            replace_in_shape(child, data)


def generate_ppt_from_template(template_path: Path, output_path: Path, data: dict):
    if not template_path.exists():
        raise HTTPException(status_code=500, detail=f"Template not found at {template_path}")

    prs = Presentation(template_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_in_shape(shape, data)

    prs.save(output_path)


def safe_filename(value: str) -> str:
    value = clean_text(value) or "Campaign"
    return re.sub(r'[\\/*?:"<>|]', "", value)


@app.post("/validate-eoc")
async def validate_eoc(payload: FileRefsPayload):
    try:
        if not payload.openaiFileIdRefs:
            raise HTTPException(status_code=400, detail="No EOC file supplied.")

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            eoc_path = await download_openai_file(payload.openaiFileIdRefs[0], tmp_dir)
            sheets = read_uploaded_eoc(eoc_path)
            result = build_mapped_values(sheets, filename=eoc_path.name)

        return JSONResponse(content={
            "status": "validated",
            "version": APP_VERSION,
            **result,
        })

    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc(),
                "version": APP_VERSION,
            },
        )


@app.post("/generate-exec-summary")
async def generate_exec_summary(payload: FileRefsPayload):
    try:
        if not payload.openaiFileIdRefs:
            raise HTTPException(status_code=400, detail="No EOC file supplied.")

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)

            eoc_path = await download_openai_file(payload.openaiFileIdRefs[0], tmp_dir)
            sheets = read_uploaded_eoc(eoc_path)

            result = build_mapped_values(sheets, filename=eoc_path.name)
            mapped_data = result["mapped_values"]

            campaign_name = safe_filename(mapped_data.get("CAMPAIGN_NAME") or "Campaign")
            filename = f"Exec Summary_PCA One Pager_{campaign_name}.pptx"
            output_path = tmp_dir / filename

            generate_ppt_from_template(TEMPLATE_PATH, output_path, mapped_data)

            if output_path.stat().st_size > MAX_RETURN_FILE_BYTES:
                raise HTTPException(status_code=413, detail="Generated PPT is over 10MB. Reduce template media size.")

            encoded = base64.b64encode(output_path.read_bytes()).decode("utf-8")

        return {
            "status": "success",
            "version": APP_VERSION,
            "summary": result,
            "openaiFileResponse": [
                {
                    "name": filename,
                    "mime_type": PPTX_MIME,
                    "content": encoded,
                }
            ],
        }

    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc(),
                "version": APP_VERSION,
            },
        )
