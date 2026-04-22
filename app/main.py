import io
import os
import re
import math
import json
import logging
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from numbers import Real
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation
from pptx.text.text import _Run

# -----------------------------------------------------------------------------
# App setup
# -----------------------------------------------------------------------------

app = FastAPI(title="PCA Automation API", version="3.0.0")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("pca-automation")

# -----------------------------------------------------------------------------
# Environment / config
# -----------------------------------------------------------------------------

RULES_PATH = os.getenv("RULES_PATH", "PCA_GPT_Rules_Master.xlsx")
DEFAULT_TEMPLATE_PATH = os.getenv("PPT_TEMPLATE_PATH", "template.pptx")
MAX_FILE_SIZE_MB = int(os.getenv("MAX_FILE_SIZE_MB", "25"))

# -----------------------------------------------------------------------------
# Normalization helpers
# -----------------------------------------------------------------------------

def norm_text(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip().lower()
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9%./:()\- ]+", "", s)
    return s.strip()


def compact_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "_", norm_text(value)).strip("_")


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip() == ""


def maybe_float(value: Any) -> Optional[float]:
    if value is None:
        return None

    if isinstance(value, (int, float, np.number)) and not pd.isna(value):
        try:
            f = float(value)
            if math.isnan(f) or math.isinf(f):
                return None
            return f
        except Exception:
            return None

    s = str(value).strip()
    if not s:
        return None

    s = s.replace(",", "")
    s = s.replace("£", "")
    s = s.replace("$", "")
    s = s.replace("€", "")
    s = s.replace("%", "")

    try:
        f = float(s)
        if math.isnan(f) or math.isinf(f):
            return None
        return f
    except Exception:
        return None


def parse_date_any(value: Any) -> Optional[datetime]:
    if value is None or value == "":
        return None

    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    if isinstance(value, datetime):
        return value

    try:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.to_pydatetime()
    except Exception:
        return None


def fmt_date_short(dt: Optional[datetime]) -> Optional[str]:
    if not dt:
        return None
    return f"{dt.day:02d} {dt.strftime('%b')}"


def fmt_date_full(dt: Optional[datetime]) -> Optional[str]:
    if not dt:
        return None
    return f"{dt.day:02d} {dt.strftime('%B %Y')}"


def quarter_from_date(dt: Optional[datetime]) -> Optional[str]:
    if not dt:
        return None
    quarter = ((dt.month - 1) // 3) + 1
    return f"Q{quarter} {dt.year}"


def format_percent_value(value: Optional[float], decimals: int = 2) -> Optional[str]:
    if value is None:
        return None

    # Rules-master-friendly sanity:
    # if value is <= 1, assume decimal rate and convert to %
    # if value is > 1, assume already percent-like only if plausible; otherwise keep as-is
    if value <= 1:
        value = value * 100

    if math.isnan(value) or math.isinf(value):
        return None
    return f"{value:.{decimals}f}%"


def format_number_value(value: Optional[float], decimals: int = 0) -> Optional[str]:
    if value is None:
        return None
    if math.isnan(value) or math.isinf(value):
        return None
    if decimals == 0:
        return f"{int(round(value)):,}"
    return f"{value:,.{decimals}f}"


def format_currency_value(value: Optional[float], currency_symbol: str = "€", decimals: int = 2) -> Optional[str]:
    if value is None:
        return None
    if math.isnan(value) or math.isinf(value):
        return None
    return f"{currency_symbol}{value:,.{decimals}f}"


def safe_jsonable(obj: Any) -> Any:
    if isinstance(obj, dict):
        return {str(k): safe_jsonable(v) for k, v in obj.items()}

    if isinstance(obj, (list, tuple, set)):
        return [safe_jsonable(v) for v in obj]

    try:
        if pd.isna(obj):
            return None
    except Exception:
        pass

    if isinstance(obj, pd.Timestamp):
        return obj.isoformat()

    if isinstance(obj, datetime):
        return obj.isoformat()

    if isinstance(obj, np.generic):
        obj = obj.item()

    if isinstance(obj, Real):
        value = float(obj)
        if math.isnan(value) or math.isinf(value):
            return None
        if value.is_integer():
            return int(value)
        return value

    if isinstance(obj, (str, bool, int)) or obj is None:
        return obj

    return str(obj)

# -----------------------------------------------------------------------------
# Rules master loading
# -----------------------------------------------------------------------------

def load_rules_master() -> Dict[str, pd.DataFrame]:
    if not os.path.exists(RULES_PATH):
        raise FileNotFoundError(f"Rules Master not found at {RULES_PATH}")

    xls = pd.ExcelFile(RULES_PATH, engine="openpyxl")
    sheets: Dict[str, pd.DataFrame] = {}

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]
        sheets[sheet] = df

    return sheets


def get_placeholder_mapping_df(rules_sheets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    df = rules_sheets.get("Placeholder Mapping")
    if df is None or df.empty:
        raise ValueError("Placeholder Mapping sheet missing or empty in Rules Master")

    work = df.copy()
    work.columns = [str(c).strip().lower() for c in work.columns]
    return work

# -----------------------------------------------------------------------------
# Alias configuration
# -----------------------------------------------------------------------------

HEADER_ALIASES: Dict[str, List[str]] = {
    "campaign": ["campaign", "campaign name", "line item"],
    "site": ["site", "publisher", "property", "app", "domain", "title"],
    "geo": ["geo", "country", "market", "region", "location", "territory"],
    "format": ["format", "ad format", "creative format", "unit type", "ad unit"],
    "date": ["date", "day", "week", "month", "reporting date", "served date"],

    "impressions": ["impressions", "imps", "served impressions", "delivered impressions"],
    "clicks": ["clicks", "click"],
    "ctr": ["ctr", "click through rate", "click-through rate"],
    "engagement_rate": ["engagement rate", "engagement %", "er", "total er", "overall er"],
    "viewability": ["mobkoi on screen", "on screen", "on-screen", "mrc viewability", "viewability", "viewable rate"],
    "video_completion_rate": ["video completion rate", "vcr", "completed view rate", "video % complete"],
    "video_completions": ["video completions", "completions", "video completes"],
    "spend": ["actual spend", "spend", "total spend", "media spend", "budget"],

    "sold_paid_units": ["sold paid units"],
    "delivered_overall_av_units": ["delivered overall av (units)"],
    "delivery_percentage_incl_av": ["delivery percentage (incl av)"],
    "delivered_av_amount_currency": ["delivered av amount (currency)"],

    "brand_client": ["client", "brand", "advertiser", "brand / client"],

    "start_date": ["start date", "campaign start", "live start", "flight start"],
    "end_date": ["end date", "campaign end", "live end", "flight end"],
}


SPECIAL_SOURCE_FIELD_ALIASES: Dict[str, List[str]] = {
    "campaign": HEADER_ALIASES["campaign"],
    "date": HEADER_ALIASES["date"] + HEADER_ALIASES["start_date"] + HEADER_ALIASES["end_date"],
    "format": HEADER_ALIASES["format"],
    "geo": HEADER_ALIASES["geo"],
    "spend": HEADER_ALIASES["spend"],
    "impressions": HEADER_ALIASES["impressions"],
    "mobkoi on screen": ["mobkoi on screen", "on screen", "on-screen", "mrc viewability", "viewability", "viewable rate"],
    "ctr": HEADER_ALIASES["ctr"],
    "engagement rate": HEADER_ALIASES["engagement_rate"],
    "video completion rate": HEADER_ALIASES["video_completion_rate"],
    "sold paid units": HEADER_ALIASES["sold_paid_units"],
    "delivered overall av (units)": HEADER_ALIASES["delivered_overall_av_units"],
    "delivery percentage (incl av)": HEADER_ALIASES["delivery_percentage_incl_av"],
    "delivered av amount (currency)": HEADER_ALIASES["delivered_av_amount_currency"],
    "brand / client": HEADER_ALIASES["brand_client"],
}

TABLE_EXPECTATIONS = {
    "campaign_kpi_summary": {
        "required_any": ["campaign", "impressions", "ctr", "engagement_rate", "video_completion_rate", "viewability", "spend"],
        "dimension": "campaign",
    },
    "campaign_delivery": {
        "required_any": ["campaign", "sold_paid_units", "delivered_overall_av_units", "delivery_percentage_incl_av", "delivered_av_amount_currency"],
        "dimension": "campaign",
    },
    "site": {
        "required_any": ["site", "ctr", "engagement_rate", "video_completion_rate", "impressions"],
        "dimension": "site",
    },
    "geo": {
        "required_any": ["geo", "ctr", "engagement_rate", "video_completion_rate", "impressions"],
        "dimension": "geo",
    },
    "format": {
        "required_any": ["format", "ctr", "engagement_rate", "video_completion_rate", "impressions"],
        "dimension": "format",
    },
    "date": {
        "required_any": ["date", "impressions"],
        "dimension": "date",
    },
}

# -----------------------------------------------------------------------------
# Candidate table dataclass
# -----------------------------------------------------------------------------

@dataclass
class CandidateTable:
    sheet_name: str
    table_type: str
    header_row_idx: int
    score: float
    matched_headers: Dict[str, str]
    n_rows: int
    n_cols: int
    preview: List[Dict[str, Any]]

# -----------------------------------------------------------------------------
# Header matching / table detection
# -----------------------------------------------------------------------------

def alias_match_score(header: str, alias: str) -> int:
    h = norm_text(header)
    a = norm_text(alias)

    if not h or not a:
        return 0
    if h == a:
        return 100
    if a in h:
        return 80
    if h in a:
        return 50
    return 0


def canonicalize_headers(raw_headers: List[Any]) -> Tuple[Dict[str, int], Dict[str, str]]:
    canonical_map: Dict[str, int] = {}
    original_header_map: Dict[str, str] = {}

    for idx, raw in enumerate(raw_headers):
        header = str(raw).strip()
        if not header:
            continue

        best_key = None
        best_score = 0

        for canonical, aliases in HEADER_ALIASES.items():
            for alias in aliases:
                score = alias_match_score(header, alias)
                if score > best_score:
                    best_score = score
                    best_key = canonical

        if best_key and best_score >= 80 and best_key not in canonical_map:
            canonical_map[best_key] = idx
            original_header_map[best_key] = header

    return canonical_map, original_header_map


def trim_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    work = work.dropna(how="all")
    work = work.dropna(axis=1, how="all")
    return work


def count_non_blank_cells(row_values: List[Any]) -> int:
    return sum(0 if is_blank(v) else 1 for v in row_values)


def build_table_from_header(raw_df: pd.DataFrame, header_idx: int) -> pd.DataFrame:
    if header_idx >= len(raw_df):
        return pd.DataFrame()

    rows = []
    blank_streak = 0

    for i in range(header_idx + 1, len(raw_df)):
        row = raw_df.iloc[i].tolist()
        if count_non_blank_cells(row) == 0:
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0
        rows.append(row)

    headers = raw_df.iloc[header_idx].tolist()
    table_df = pd.DataFrame(rows, columns=headers)
    return trim_dataframe(table_df)


def score_table_type(canonical_headers: Dict[str, int], table_type: str, table_df: pd.DataFrame) -> float:
    rule = TABLE_EXPECTATIONS[table_type]
    score = 0.0

    hits = 0
    for key in rule["required_any"]:
        if key in canonical_headers:
            hits += 1
            score += 18

    if hits >= 3:
        score += 20

    dim = rule["dimension"]
    if dim in canonical_headers:
        score += 30

    if len(table_df) >= 2:
        score += 5
    if len(table_df) >= 5:
        score += 5

    # strong preference for rate columns in KPI/site tables
    if table_type in {"campaign_kpi_summary", "site", "geo", "format"}:
        for k in ["ctr", "engagement_rate", "video_completion_rate"]:
            if k in canonical_headers:
                score += 10

    if table_type == "campaign_delivery":
        for k in ["sold_paid_units", "delivered_overall_av_units", "delivery_percentage_incl_av", "delivered_av_amount_currency"]:
            if k in canonical_headers:
                score += 12

    return score


def find_candidate_tables(xls: pd.ExcelFile) -> List[CandidateTable]:
    candidates: List[CandidateTable] = []

    for sheet_name in xls.sheet_names:
        try:
            raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        except Exception as e:
            logger.warning("Could not read sheet %s: %s", sheet_name, e)
            continue

        raw_df = trim_dataframe(raw_df)
        if raw_df.empty:
            continue

        for row_idx in range(min(len(raw_df), 80)):
            row_values = raw_df.iloc[row_idx].tolist()

            if count_non_blank_cells(row_values) < 2:
                continue

            canonical_headers, original_header_map = canonicalize_headers(row_values)
            if not canonical_headers:
                continue

            table_df = build_table_from_header(raw_df, row_idx)
            if table_df.empty:
                continue

            preview = safe_jsonable(table_df.head(3).fillna("").to_dict(orient="records"))

            for table_type in TABLE_EXPECTATIONS.keys():
                score = score_table_type(canonical_headers, table_type, table_df)
                if score >= 50:
                    candidates.append(
                        CandidateTable(
                            sheet_name=sheet_name,
                            table_type=table_type,
                            header_row_idx=row_idx,
                            score=score,
                            matched_headers=original_header_map,
                            n_rows=len(table_df),
                            n_cols=len(table_df.columns),
                            preview=preview,
                        )
                    )

    candidates.sort(key=lambda c: c.score, reverse=True)
    return candidates


def standardize_table(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    headers = list(work.columns)
    canonical_map, _ = canonicalize_headers(headers)

    rename_map = {}
    for canonical, idx in canonical_map.items():
        rename_map[headers[idx]] = canonical

    work = work.rename(columns=rename_map)
    return work


def remove_duplicate_header_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    work = df.copy()
    col0 = str(work.columns[0]).strip().lower()

    def is_repeated_header(row: pd.Series) -> bool:
        vals = [norm_text(v) for v in row.tolist()]
        header_vals = [norm_text(c) for c in work.columns.tolist()]
        matches = 0
        for a, b in zip(vals[: min(8, len(vals))], header_vals[: min(8, len(header_vals))]):
            if a and b and a == b:
                matches += 1
        return matches >= 3 or (vals and vals[0] == col0)

    mask = work.apply(lambda row: not is_repeated_header(row), axis=1)
    work = work[mask]
    return work.reset_index(drop=True)


def coerce_types(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()

    for col in work.columns:
        lname = norm_text(col)
        if any(k in lname for k in ["ctr", "rate", "viewability", "spend", "impressions", "clicks", "completions", "units", "amount"]):
            try:
                work[col] = work[col].apply(lambda x: maybe_float(x) if not isinstance(x, datetime) else x)
            except Exception:
                pass

        if "date" in lname:
            try:
                work[col] = work[col].apply(parse_date_any)
            except Exception:
                pass

    return work


def select_best_tables(candidates: List[CandidateTable], xls: pd.ExcelFile) -> Dict[str, Dict[str, Any]]:
    selected: Dict[str, Dict[str, Any]] = {}

    for table_type in TABLE_EXPECTATIONS.keys():
        type_candidates = [c for c in candidates if c.table_type == table_type]
        if not type_candidates:
            continue

        best = type_candidates[0]
        raw_df = pd.read_excel(xls, sheet_name=best.sheet_name, header=None)
        table_df = build_table_from_header(trim_dataframe(raw_df), best.header_row_idx)
        table_df = standardize_table(table_df)
        table_df = remove_duplicate_header_rows(table_df)
        table_df = coerce_types(table_df)

        selected[table_type] = {
            "candidate": best,
            "df": table_df,
        }

    return selected

# -----------------------------------------------------------------------------
# Workbook metadata / title scanning
# -----------------------------------------------------------------------------

def scan_workbook_text_cells(xls: pd.ExcelFile, max_rows: int = 25, max_cols: int = 8) -> List[str]:
    texts: List[str] = []

    for sheet_name in xls.sheet_names:
        try:
            raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        except Exception:
            continue

        raw = raw.iloc[:max_rows, :max_cols]
        for _, row in raw.iterrows():
            for val in row.tolist():
                if not is_blank(val):
                    texts.append(str(val).strip())

    return texts


def extract_report_date_from_texts(texts: List[str]) -> Optional[datetime]:
    pattern = re.compile(r"report date[: ]+(\d{4}-\d{2}-\d{2})", re.I)
    for txt in texts:
        m = pattern.search(txt)
        if m:
            return parse_date_any(m.group(1))
    return None


def infer_brand_from_filename(filename: str) -> Optional[str]:
    if not filename:
        return None
    base = os.path.basename(filename)
    base = os.path.splitext(base)[0]
    parts = re.split(r"\s*[-_]\s*", base)
    if parts:
        candidate = parts[0].strip()
        if candidate and len(candidate) <= 40:
            return candidate
    return None

# -----------------------------------------------------------------------------
# Rules execution helpers
# -----------------------------------------------------------------------------

def map_table_section_to_internal(section: str) -> str:
    s = norm_text(section)

    if "campaign kpi summary" in s:
        return "campaign_kpi_summary"
    if "campaign delivery" in s:
        return "campaign_delivery"
    if s == "site":
        return "site"
    if s == "geo":
        return "geo"
    if s == "format":
        return "format"
    if "date / metadata" in s:
        return "date_meta"
    if "brand / client" in s:
        return "brand_client"

    return compact_key(section)


def find_best_column_name(df: pd.DataFrame, source_field_name: str) -> Optional[str]:
    target = norm_text(source_field_name)
    aliases = SPECIAL_SOURCE_FIELD_ALIASES.get(target, [target])

    best_col = None
    best_score = 0

    for col in df.columns:
        col_text = norm_text(col)
        for alias in aliases:
            score = alias_match_score(col_text, alias)
            if score > best_score:
                best_score = score
                best_col = col

    return best_col if best_score >= 70 else None


def get_first_non_empty_value(series: pd.Series) -> Any:
    for v in series.tolist():
        if not is_blank(v):
            return v
    return None


def resolve_brand_client(
    xls: pd.ExcelFile,
    selected_tables: Dict[str, Dict[str, Any]],
    filename: str
) -> Optional[str]:
    # Priority 1/2: explicit field in report / campaign table
    for table_name in ["campaign_kpi_summary", "campaign_delivery"]:
        if table_name in selected_tables:
            df = selected_tables[table_name]["df"]
            col = find_best_column_name(df, "brand / client")
            if col:
                val = get_first_non_empty_value(df[col])
                if val:
                    return str(val).strip()

    # Priority 3: workbook metadata / title text
    texts = scan_workbook_text_cells(xls)
    for txt in texts:
        low = norm_text(txt)
        # basic but safe: ignore boilerplate report strings
        if len(low) < 3:
            continue
        if "weekly report" in low or "report date" in low:
            continue
        # if a title looks like "Prada Eyewear FW25"
        if len(txt.strip()) <= 40 and txt.strip().istitle():
            return txt.strip()

    # Priority 4: filename
    brand = infer_brand_from_filename(filename)
    if brand:
        return brand

    # Priority 5/6: fallback
    return None


def resolve_dates(
    xls: pd.ExcelFile,
    selected_tables: Dict[str, Dict[str, Any]],
    filename: str
) -> Tuple[Optional[datetime], Optional[datetime], Optional[datetime]]:
    # 1) Date table first
    if "date" in selected_tables:
        df = selected_tables["date"]["df"]
        col = find_best_column_name(df, "date")
        if col:
            vals = [parse_date_any(v) for v in df[col].tolist()]
            vals = [v for v in vals if v]
            if vals:
                return min(vals), max(vals), max(vals)

    # 2) Metadata / title cells
    texts = scan_workbook_text_cells(xls)
    report_date = extract_report_date_from_texts(texts)
    if report_date:
        return None, None, report_date

    # 3) Filename fallback: YYYY-MM-DD -> period end date, minus 6 days
    m = re.search(r"(20\d{2}-\d{2}-\d{2})", filename or "")
    if m:
        end_date = parse_date_any(m.group(1))
        if end_date:
            return end_date - timedelta(days=6), end_date, end_date

    return None, None, None


def extract_direct_value(df: pd.DataFrame, source_field_name: str) -> Any:
    col = find_best_column_name(df, source_field_name)
    if not col:
        return None

    series = df[col].dropna()
    if series.empty:
        return None

    return series.iloc[0]


def extract_join_non_empty_rows(df: pd.DataFrame, source_field_name: str) -> Optional[str]:
    col = find_best_column_name(df, source_field_name)
    if not col:
        return None

    vals = []
    for v in df[col].tolist():
        if is_blank(v):
            continue
        s = str(v).strip()
        low = norm_text(s)
        if low in {"campaign", "site", "format", "geo", "country", "market", "region"}:
            continue
        if s not in vals:
            vals.append(s)

    if not vals:
        return None

    return ", ".join(vals)


def execute_rule_for_placeholder(
    rule: pd.Series,
    xls: pd.ExcelFile,
    selected_tables: Dict[str, Dict[str, Any]],
    filename: str
) -> Any:
    placeholder = str(rule.get("placeholder", "")).strip()
    table_section = str(rule.get("table / section", "")).strip()
    source_field_name = str(rule.get("source field name", "")).strip()
    logic = str(rule.get("logic", "")).strip().lower()
    fmt = str(rule.get("format", "")).strip().lower()

    internal_table = map_table_section_to_internal(table_section)

    # Special: Brand / Client
    if internal_table == "brand_client":
        value = resolve_brand_client(xls, selected_tables, filename)
        return value

    # Special: Date / Metadata
    if internal_table == "date_meta":
        start_date, end_date, fallback_date = resolve_dates(xls, selected_tables, filename)

        if "quarter" in logic:
            base = end_date or fallback_date or start_date
            return quarter_from_date(base)

        if "short format" in logic:
            if start_date and end_date:
                return f"{fmt_date_short(start_date)} - {fmt_date_short(end_date)}"
            if fallback_date:
                return fmt_date_short(fallback_date)
            return None

        if "full format" in logic:
            if start_date and end_date:
                return f"{fmt_date_full(start_date)} - {fmt_date_full(end_date)}"
            if fallback_date:
                return fmt_date_full(fallback_date)
            return None

        return None

    table_entry = selected_tables.get(internal_table)
    if not table_entry:
        return None

    df = table_entry["df"]

    # Top titles patterns handled separately later
    if "#" in placeholder or "*" in placeholder:
        return None

    # Direct / N/A if missing / direct abs if needed
    if "join non-empty rows" in logic:
        return extract_join_non_empty_rows(df, source_field_name)

    value = extract_direct_value(df, source_field_name)

    if value is None:
        return None

    # formatting driven after extraction
    if isinstance(value, (int, float, np.number)):
        value = float(value)

    if "abs if needed" in logic and isinstance(value, (int, float)):
        value = abs(value)

    if fmt == "%":
        return format_percent_value(maybe_float(value), 2)
    if fmt == "number":
        return format_number_value(maybe_float(value), 0)
    if fmt == "currency":
        return format_currency_value(maybe_float(value), "€", 2)

    return str(value).strip() if not isinstance(value, (int, float)) else value


def build_top_titles_placeholders(
    selected_tables: Dict[str, Dict[str, Any]],
    top_rank: int = 5
) -> Dict[str, Any]:
    results: Dict[str, Any] = {}

    if "site" not in selected_tables:
        return results

    df = selected_tables["site"]["df"].copy()

    site_col = find_best_column_name(df, "site")
    ctr_col = find_best_column_name(df, "ctr")
    er_col = find_best_column_name(df, "engagement rate")
    vcr_col = find_best_column_name(df, "video completion rate")

    if not site_col:
        return results

    def build_metric(metric_col: Optional[str], metric_key: str) -> None:
        if not metric_col:
            return

        work = df[[site_col, metric_col]].copy()
        work = work.dropna(subset=[site_col, metric_col])

        # remove header-like rows
        work = work[work[site_col].astype(str).str.strip().str.lower() != "site"]
        work = work[~work[site_col].astype(str).str.strip().str.lower().isin(["campaign", "format", "geo", "country", "market", "region"])]

        if work.empty:
            return

        work[metric_col] = work[metric_col].apply(maybe_float)
        work = work.dropna(subset=[metric_col])

        if work.empty:
            return

        work = work.sort_values(metric_col, ascending=False).head(top_rank)

        rows = work.to_dict(orient="records")
        for idx, row in enumerate(rows, start=1):
            name_placeholder = f"{{{{TOP_TITLES_{metric_key}_{idx}_NAME}}}}"
            value_placeholder = f"{{{{TOP_TITLES_{metric_key}_{idx}_VALUE}}}}"
            results[name_placeholder] = str(row[site_col]).strip()
            results[value_placeholder] = format_percent_value(maybe_float(row[metric_col]), 2)

    build_metric(ctr_col, "CTR")
    build_metric(er_col, "ER")
    build_metric(vcr_col, "VCR")

    return results

# -----------------------------------------------------------------------------
# Main mapping builder
# -----------------------------------------------------------------------------

def build_mapped_values(xls: pd.ExcelFile, filename: str = "") -> Dict[str, Any]:
    rules_sheets = load_rules_master()
    mapping_df = get_placeholder_mapping_df(rules_sheets)

    candidates = find_candidate_tables(xls)
    selected_tables = select_best_tables(candidates, xls)

    mapped_values: Dict[str, Any] = {}

    for _, rule in mapping_df.iterrows():
        placeholder = str(rule.get("placeholder", "")).strip()
        if not placeholder:
            continue

        if "#_*" in placeholder:
            # handled later
            continue

        value = execute_rule_for_placeholder(rule, xls, selected_tables, filename)
        mapped_values[placeholder] = value

    # Top titles based on rules doc safe fallback to 5
    mapped_values.update(build_top_titles_placeholders(selected_tables, top_rank=5))

    # Validation
    required_core = [
        "{{CAMPAIGN_NAME}}",
        "{{LIVE_DATES_SHORT}}",
        "{{PERFORMANCE_CTR}}",
        "{{PERFORMANCE_ENGAGEMENT_RATE}}",
        "{{PERFORMANCE_VCR}}",
    ]
    missing_required = [k for k in required_core if not mapped_values.get(k)]

    warnings = []
    if "geo" not in selected_tables and not mapped_values.get("{{CAMPAIGN_MARKETS}}"):
        warnings.append("No geo table found; markets unresolved")
    if "campaign_delivery" not in selected_tables:
        warnings.append("No separate delivery table found")
    if "site" not in selected_tables:
        warnings.append("No site table found")

    diagnostics = {
        "candidate_tables": [asdict(c) for c in candidates[:20]],
        "selected_tables": {
            key: {
                "candidate": asdict(val["candidate"]),
                "columns": [str(c) for c in val["df"].columns.tolist()],
                "preview": safe_jsonable(val["df"].head(5).fillna("").to_dict(orient="records")),
            }
            for key, val in selected_tables.items()
        },
    }

    return safe_jsonable(
        {
            "mapped_values": mapped_values,
            "validation": {
                "is_valid": len(missing_required) == 0,
                "missing_required": missing_required,
                "warnings": warnings,
            },
            "diagnostics": diagnostics,
        }
    )

# -----------------------------------------------------------------------------
# PPT placeholder replacement
# -----------------------------------------------------------------------------

def replace_text_in_run(run: _Run, replacements: Dict[str, str]) -> None:
    text = run.text
    for placeholder, value in replacements.items():
        text = text.replace(placeholder, value)
    run.text = text


def replace_text_in_text_frame(text_frame, replacements: Dict[str, str]) -> None:
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            replace_text_in_run(run, replacements)

        combined = "".join(run.text for run in paragraph.runs)
        new_text = combined
        for placeholder, value in replacements.items():
            new_text = new_text.replace(placeholder, value)

        if new_text != combined and paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""


def replace_in_shape(shape, replacements: Dict[str, str]) -> None:
    if hasattr(shape, "text_frame") and shape.has_text_frame:
        replace_text_in_text_frame(shape.text_frame, replacements)

    if shape.shape_type == 6:
        for subshape in shape.shapes:
            replace_in_shape(subshape, replacements)

    if hasattr(shape, "table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_text_in_text_frame(cell.text_frame, replacements)


def fill_ppt_template(template_bytes: bytes, mapped_values: Dict[str, Any]) -> bytes:
    prs = Presentation(io.BytesIO(template_bytes))

    replacements = {k: "" if v is None else str(v) for k, v in mapped_values.items()}

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_in_shape(shape, replacements)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# -----------------------------------------------------------------------------
# File helpers
# -----------------------------------------------------------------------------

async def read_upload_bytes(file: UploadFile) -> bytes:
    content = await file.read()
    size_mb = len(content) / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        raise HTTPException(status_code=400, detail=f"File too large. Max {MAX_FILE_SIZE_MB}MB")
    return content


def open_excel_from_bytes(content: bytes) -> pd.ExcelFile:
    try:
        return pd.ExcelFile(io.BytesIO(content), engine="openpyxl")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid Excel file: {e}")

# -----------------------------------------------------------------------------
# Routes
# -----------------------------------------------------------------------------

@app.get("/health")
def health() -> Dict[str, str]:
    return {"status": "ok"}


@app.post("/validate-eoc")
async def validate_eoc(file: UploadFile = File(...)) -> JSONResponse:
    try:
        if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            raise HTTPException(status_code=400, detail="Please upload an Excel file")

        content = await read_upload_bytes(file)
        xls = open_excel_from_bytes(content)
        result = build_mapped_values(xls, filename=file.filename)

        return JSONResponse(content=safe_jsonable(result))

    except HTTPException:
        raise
    except Exception as e:
        logger.exception("validate_eoc failed")
        raise HTTPException(status_code=500, detail=f"validate_eoc failed: {str(e)}")


@app.post("/generate-ppt")
async def generate_ppt(
    file: UploadFile = File(...),
    use_server_template: bool = Form(True),
    template_file: Optional[UploadFile] = File(None),
) -> StreamingResponse:
    try:
        if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            raise HTTPException(status_code=400, detail="Please upload an Excel file")

        eoc_content = await read_upload_bytes(file)
        xls = open_excel_from_bytes(eoc_content)
        result = build_mapped_values(xls, filename=file.filename)

        if use_server_template:
            if not os.path.exists(DEFAULT_TEMPLATE_PATH):
                raise HTTPException(status_code=500, detail=f"Template not found at {DEFAULT_TEMPLATE_PATH}")
            with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
                template_bytes = f.read()
        else:
            if template_file is None:
                raise HTTPException(status_code=400, detail="No template file uploaded")
            if not template_file.filename.lower().endswith(".pptx"):
                raise HTTPException(status_code=400, detail="Template must be a .pptx file")
            template_bytes = await read_upload_bytes(template_file)

        ppt_bytes = fill_ppt_template(template_bytes, result["mapped_values"])

        return StreamingResponse(
            io.BytesIO(ppt_bytes),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": 'attachment; filename="exec_summary_output.pptx"',
                "X-Mapped-Values": json.dumps(safe_jsonable(result["mapped_values"])),
            },
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.exception("generate_ppt failed")
        raise HTTPException(status_code=500, detail=f"generate_ppt failed: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.getenv("PORT", "8000")), reload=True)
