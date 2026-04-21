import io
import os
import re
import math
import json
import logging
import numpy as np
from numbers import Real
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation
from pptx.text.text import _Run

app = FastAPI(title="PCA Automation API", version="2.0.1")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("pca-automation")

DEFAULT_TEMPLATE_PATH = os.getenv("PPT_TEMPLATE_PATH", "template.pptx")
MAX_FILE_SIZE_MB = int(os.getenv("MAX_FILE_SIZE_MB", "25"))


def norm_text(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip().lower()
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    s = s.replace("%", " percent ")
    s = re.sub(r"[^a-z0-9./:()\- ]+", "", s)
    return s.strip()


def compact_key(value: Any) -> str:
    s = norm_text(value)
    s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
    return s


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return str(value).strip() == ""


def maybe_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    s = s.replace(",", "")
    s = s.replace("%", "")
    s = s.replace("£", "")
    s = s.replace("$", "")
    try:
        return float(s)
    except Exception:
        return None


def fmt_percent(value: Optional[float], decimals: int = 2) -> Optional[str]:
    if value is None:
        return None
    if math.isnan(value) or math.isinf(value):
        return None
    return f"{value:.{decimals}f}%"


def fmt_int(value: Optional[float]) -> Optional[str]:
    if value is None:
        return None
    if math.isnan(value) or math.isinf(value):
        return None
    return f"{int(round(value)):,}"


def fmt_num(value: Optional[float], decimals: int = 2) -> Optional[str]:
    if value is None:
        return None
    if math.isnan(value) or math.isinf(value):
        return None
    return f"{value:,.{decimals}f}"


def parse_date_any(value: Any) -> Optional[datetime]:
    if value is None or value == "":
        return None
    try:
        if isinstance(value, datetime):
            return value
        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime()
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.to_pydatetime()
    except Exception:
        return None


def fmt_date_short(dt: Optional[datetime]) -> Optional[str]:
    if not dt:
        return None
    return f"{dt.day} {dt.strftime('%b')}"


def fmt_date_full(dt: Optional[datetime]) -> Optional[str]:
    if not dt:
        return None
    return f"{dt.day} {dt.strftime('%B %Y')}"


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
        if float(value).is_integer():
            return int(value)
        return value

    if isinstance(obj, (str, bool, int)) or obj is None:
        return obj

    return str(obj)


HEADER_ALIASES: Dict[str, List[str]] = {
    "campaign": [
        "campaign", "campaign name", "campaign title", "line item", "package", "placement"
    ],
    "site": [
        "site", "publisher", "property", "app", "domain", "title", "site/app", "inventory source"
    ],
    "geo": [
        "geo", "country", "market", "region", "location", "territory"
    ],
    "format": [
        "format", "ad format", "creative format", "unit type", "device format", "ad unit"
    ],
    "date": [
        "date", "day", "week", "month", "reporting date", "served date"
    ],
    "impressions": [
        "impressions", "imps", "delivered impressions", "served impressions"
    ],
    "clicks": [
        "clicks", "click"
    ],
    "ctr": [
        "ctr", "click through rate", "click-through rate"
    ],
    "viewability": [
        "viewability", "viewable rate", "viewability rate"
    ],
    "engagement_rate": [
        "engagement rate", "er", "eng rate", "interaction rate"
    ],
    "vcr": [
        "vcr", "video completion rate", "completion rate", "video completion"
    ],
    "completes": [
        "completes", "completions", "video completes"
    ],
    "planned_impressions": [
        "planned impressions", "booked impressions", "goal impressions", "target impressions"
    ],
    "delivered_impressions": [
        "delivered impressions", "served impressions", "actual impressions"
    ],
    "ad_value": [
        "ad value", "media value", "av", "earned media value"
    ],
    "spend": [
        "spend", "cost", "investment", "media spend", "budget"
    ],
    "start_date": [
        "start date", "campaign start", "live start", "flight start"
    ],
    "end_date": [
        "end date", "campaign end", "live end", "flight end"
    ],
}

TABLE_TYPE_RULES = {
    "campaign": {
        "required_any": ["campaign", "start_date", "end_date"],
        "nice_to_have": ["geo", "format", "impressions"],
    },
    "site": {
        "required_any": ["site"],
        "nice_to_have": ["impressions", "clicks", "ctr", "viewability", "engagement_rate", "vcr"],
    },
    "geo": {
        "required_any": ["geo"],
        "nice_to_have": ["impressions", "clicks", "ctr", "viewability", "engagement_rate", "vcr"],
    },
    "format": {
        "required_any": ["format"],
        "nice_to_have": ["impressions", "clicks", "ctr", "viewability", "engagement_rate", "vcr"],
    },
    "date": {
        "required_any": ["date"],
        "nice_to_have": ["impressions", "clicks", "ctr", "viewability", "engagement_rate", "vcr"],
    },
    "kpi": {
        "required_any": ["ctr", "viewability", "engagement_rate", "vcr"],
        "nice_to_have": ["clicks", "impressions", "completes"],
    },
    "delivery": {
        "required_any": ["planned_impressions", "delivered_impressions", "impressions", "ad_value", "spend"],
        "nice_to_have": ["ctr", "viewability"],
    },
}

PLACEHOLDER_RULES = {
    "CAMPAIGN_NAME": "campaign_name",
    "LIVE_DATES_SHORT": "live_dates_short",
    "LIVE_DATES_FULL": "live_dates_full",
    "PERFORMANCE_CTR": "performance_ctr",
    "PERFORMANCE_VIEWABILITY": "performance_viewability",
    "PERFORMANCE_ENGAGEMENT_RATE": "performance_engagement_rate",
    "PERFORMANCE_VCR": "performance_vcr",
    "TOTAL_IMPRESSIONS": "total_impressions",
    "PLANNED_IMPRESSIONS": "planned_impressions",
    "DELIVERED_IMPRESSIONS": "delivered_impressions",
    "AD_VALUE": "ad_value",
    "TOP_SITE_1_NAME": "top_site_1_name",
    "TOP_SITE_1_CTR": "top_site_1_ctr",
    "TOP_SITE_2_NAME": "top_site_2_name",
    "TOP_SITE_2_CTR": "top_site_2_ctr",
    "TOP_GEO_1_NAME": "top_geo_1_name",
    "TOP_FORMAT_1_NAME": "top_format_1_name",
}


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


def alias_match_score(header: str, alias: str) -> int:
    h = norm_text(header)
    a = norm_text(alias)
    if not h or not a:
        return 0
    if h == a:
        return 100
    if a in h:
        return 70
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

        if best_key and best_score >= 70 and best_key not in canonical_map:
            canonical_map[best_key] = idx
            original_header_map[best_key] = header

    return canonical_map, original_header_map


def count_non_blank_cells(row_values: List[Any]) -> int:
    return sum(0 if is_blank(v) else 1 for v in row_values)


def trim_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    return df


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


def score_table_type(canonical_headers: Dict[str, int], table_type: str) -> float:
    rules = TABLE_TYPE_RULES[table_type]
    score = 0.0

    for key in rules["required_any"]:
        if key in canonical_headers:
            score += 30

    for key in rules["nice_to_have"]:
        if key in canonical_headers:
            score += 10

    score += min(len(canonical_headers), 8) * 2

    if table_type in {"site", "geo", "format", "date"} and table_type in canonical_headers:
        score += 20

    if table_type == "kpi":
        metric_keys = {"ctr", "viewability", "engagement_rate", "vcr", "clicks", "impressions", "completes"}
        score += len(metric_keys.intersection(canonical_headers.keys())) * 4

    if table_type == "delivery":
        if "planned_impressions" in canonical_headers and "delivered_impressions" in canonical_headers:
            score += 35
        if "ad_value" in canonical_headers:
            score += 10

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

        for row_idx in range(min(len(raw_df), 60)):
            row_values = raw_df.iloc[row_idx].tolist()
            if count_non_blank_cells(row_values) < 2:
                continue

            canonical_headers, original_header_map = canonicalize_headers(row_values)
            if not canonical_headers:
                continue

            table_df = build_table_from_header(raw_df, row_idx)
            if table_df.empty or len(table_df) < 1:
                continue

            preview = safe_jsonable(table_df.head(3).fillna("").to_dict(orient="records"))

            for table_type in TABLE_TYPE_RULES.keys():
                score = score_table_type(canonical_headers, table_type)

                if len(table_df.columns) < 2:
                    score -= 20
                if len(table_df) >= 3:
                    score += 5
                if len(canonical_headers) >= 3:
                    score += 5

                if score >= 45:
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


def select_best_tables(candidates: List[CandidateTable], xls: pd.ExcelFile) -> Dict[str, Dict[str, Any]]:
    selected: Dict[str, Dict[str, Any]] = {}

    for table_type in TABLE_TYPE_RULES.keys():
        type_candidates = [c for c in candidates if c.table_type == table_type]
        if not type_candidates:
            continue

        best = type_candidates[0]
        raw_df = pd.read_excel(xls, sheet_name=best.sheet_name, header=None)
        table_df = build_table_from_header(trim_dataframe(raw_df), best.header_row_idx)

        selected[table_type] = {
            "candidate": best,
            "df": table_df,
        }

    return selected


def standardize_table(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    headers = list(df.columns)
    canonical_map, _ = canonicalize_headers(headers)

    rename_map = {}
    for canonical_key, col_idx in canonical_map.items():
        rename_map[headers[col_idx]] = canonical_key

    df = df.rename(columns=rename_map)
    return df


def coerce_metric_columns(df: pd.DataFrame) -> pd.DataFrame:
    metric_cols = [
        "impressions", "clicks", "ctr", "viewability", "engagement_rate", "vcr",
        "completes", "planned_impressions", "delivered_impressions", "ad_value", "spend"
    ]
    df = df.copy()

    for col in metric_cols:
        if col in df.columns:
            df[col] = df[col].apply(maybe_float)

    for date_col in ["start_date", "end_date", "date"]:
        if date_col in df.columns:
            df[date_col] = df[date_col].apply(parse_date_any)

    return df


LABEL_VALUE_ALIASES = {
    "campaign_name": ["campaign", "campaign name"],
    "start_date": ["start date", "campaign start", "live start"],
    "end_date": ["end date", "campaign end", "live end"],
}


def extract_label_value_metadata(xls: pd.ExcelFile) -> Dict[str, Any]:
    found: Dict[str, Any] = {}

    for sheet_name in xls.sheet_names:
        try:
            raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        except Exception:
            continue

        raw_df = trim_dataframe(raw_df)
        if raw_df.empty:
            continue

        for r in range(min(len(raw_df), 80)):
            row = raw_df.iloc[r].tolist()
            for i in range(len(row) - 1):
                left = norm_text(row[i])
                right = row[i + 1]

                for target_key, aliases in LABEL_VALUE_ALIASES.items():
                    if target_key in found and found[target_key] not in [None, ""]:
                        continue
                    if any(left == norm_text(alias) for alias in aliases):
                        found[target_key] = right

    if "start_date" in found:
        found["start_date"] = parse_date_any(found["start_date"])
    if "end_date" in found:
        found["end_date"] = parse_date_any(found["end_date"])

    return found


def find_total_row(df: pd.DataFrame, label_columns: List[str]) -> Optional[pd.Series]:
    total_terms = {"total", "grand total", "overall", "all", "campaign total"}

    for _, row in df.iterrows():
        for col in label_columns:
            if col in df.columns:
                val = norm_text(row.get(col))
                if val in total_terms:
                    return row
    return None


def choose_metric_from_tables(
    selected_tables: Dict[str, Dict[str, Any]],
    metric_key: str
) -> Optional[float]:
    if "kpi" in selected_tables:
        kpi_df = coerce_metric_columns(standardize_table(selected_tables["kpi"]["df"]))
        if metric_key in kpi_df.columns:
            series = kpi_df[metric_key].dropna()
            if not series.empty:
                value = float(series.iloc[-1])
                if math.isnan(value) or math.isinf(value):
                    return None
                return value

    if "delivery" in selected_tables:
        del_df = coerce_metric_columns(standardize_table(selected_tables["delivery"]["df"]))
        if metric_key in del_df.columns:
            total_row = find_total_row(del_df, ["campaign", "site", "geo", "format"])
            if total_row is not None and pd.notna(total_row.get(metric_key)):
                value = float(total_row.get(metric_key))
                if math.isnan(value) or math.isinf(value):
                    return None
                return value
            series = del_df[metric_key].dropna()
            if not series.empty:
                value = float(series.iloc[-1])
                if math.isnan(value) or math.isinf(value):
                    return None
                return value

    for table_type in ["site", "geo", "format", "date", "campaign"]:
        if table_type not in selected_tables:
            continue
        df = coerce_metric_columns(standardize_table(selected_tables[table_type]["df"]))

        if metric_key not in df.columns:
            continue

        total_row = find_total_row(df, ["campaign", "site", "geo", "format", "date"])
        if total_row is not None and pd.notna(total_row.get(metric_key)):
            value = float(total_row.get(metric_key))
            if math.isnan(value) or math.isinf(value):
                return None
            return value

        series = df[metric_key].dropna()
        if not series.empty:
            if metric_key in {"ctr", "viewability", "engagement_rate", "vcr"}:
                value = float(series.median())
            else:
                value = float(series.sum())

            if math.isnan(value) or math.isinf(value):
                return None
            return value

    return None


def choose_campaign_name(meta: Dict[str, Any], selected_tables: Dict[str, Dict[str, Any]]) -> Optional[str]:
    if meta.get("campaign_name"):
        return str(meta["campaign_name"]).strip()

    for table_type in ["campaign", "delivery", "site"]:
        if table_type not in selected_tables:
            continue
        df = standardize_table(selected_tables[table_type]["df"])
        if "campaign" in df.columns:
            vals = [str(v).strip() for v in df["campaign"].dropna().tolist() if str(v).strip()]
            vals = [v for v in vals if norm_text(v) not in {"total", "grand total", "overall"}]
            if vals:
                return vals[0]
    return None


def choose_live_dates(meta: Dict[str, Any], selected_tables: Dict[str, Dict[str, Any]]) -> Tuple[Optional[datetime], Optional[datetime]]:
    start_date = meta.get("start_date")
    end_date = meta.get("end_date")

    if start_date and end_date:
        return start_date, end_date

    if "campaign" in selected_tables:
        df = coerce_metric_columns(standardize_table(selected_tables["campaign"]["df"]))
        if not start_date and "start_date" in df.columns:
            vals = [v for v in df["start_date"].dropna().tolist() if v]
            if vals:
                start_date = min(vals)
        if not end_date and "end_date" in df.columns:
            vals = [v for v in df["end_date"].dropna().tolist() if v]
            if vals:
                end_date = max(vals)

    if "date" in selected_tables:
        df = coerce_metric_columns(standardize_table(selected_tables["date"]["df"]))
        if "date" in df.columns:
            vals = [v for v in df["date"].dropna().tolist() if v]
            if vals:
                if not start_date:
                    start_date = min(vals)
                if not end_date:
                    end_date = max(vals)

    return start_date, end_date


def extract_top_rows(
    selected_tables: Dict[str, Dict[str, Any]],
    table_type: str,
    sort_metric: str = "ctr",
    top_n: int = 2
) -> List[Dict[str, Any]]:
    if table_type not in selected_tables:
        return []

    df = coerce_metric_columns(standardize_table(selected_tables[table_type]["df"]))

    if table_type not in df.columns or sort_metric not in df.columns:
        return []

    label_col = table_type
    work = df[[label_col, sort_metric]].copy()
    work = work.dropna(subset=[label_col, sort_metric])
    work[label_col] = work[label_col].astype(str).str.strip()
    work = work[~work[label_col].str.lower().isin(["total", "grand total", "overall", "all", "campaign total"])]

    if work.empty:
        return []

    work = work.sort_values(sort_metric, ascending=False).head(top_n)
    return safe_jsonable(work.to_dict(orient="records"))


def validate_mapped_values(mapped: Dict[str, Any]) -> Dict[str, Any]:
    required_core = [
        "campaign_name",
        "live_dates_short",
        "performance_ctr",
        "performance_viewability",
        "performance_engagement_rate",
        "performance_vcr",
    ]

    missing_required = [k for k in required_core if not mapped.get(k)]

    warnings = []
    if not mapped.get("planned_impressions") and not mapped.get("delivered_impressions"):
        warnings.append("No delivery totals found")
    if not mapped.get("top_site_1_name"):
        warnings.append("No top site ranking found")
    if not mapped.get("top_geo_1_name"):
        warnings.append("No geo ranking found")
    if not mapped.get("top_format_1_name"):
        warnings.append("No format ranking found")

    return {
        "is_valid": len(missing_required) == 0,
        "missing_required": missing_required,
        "warnings": warnings,
    }


def build_mapped_values(xls: pd.ExcelFile) -> Dict[str, Any]:
    candidates = find_candidate_tables(xls)
    selected_tables = select_best_tables(candidates, xls)
    meta = extract_label_value_metadata(xls)

    campaign_name = choose_campaign_name(meta, selected_tables)
    start_date, end_date = choose_live_dates(meta, selected_tables)

    ctr = choose_metric_from_tables(selected_tables, "ctr")
    viewability = choose_metric_from_tables(selected_tables, "viewability")
    engagement_rate = choose_metric_from_tables(selected_tables, "engagement_rate")
    vcr = choose_metric_from_tables(selected_tables, "vcr")

    planned_impressions = choose_metric_from_tables(selected_tables, "planned_impressions")
    delivered_impressions = choose_metric_from_tables(selected_tables, "delivered_impressions")
    total_impressions = choose_metric_from_tables(selected_tables, "impressions")
    ad_value = choose_metric_from_tables(selected_tables, "ad_value")

    top_sites = extract_top_rows(selected_tables, "site", sort_metric="ctr", top_n=2)
    top_geos = extract_top_rows(selected_tables, "geo", sort_metric="ctr", top_n=1)
    top_formats = extract_top_rows(selected_tables, "format", sort_metric="ctr", top_n=1)

    mapped = {
        "campaign_name": campaign_name,
        "live_dates_short": (
            f"{fmt_date_short(start_date)} - {fmt_date_short(end_date)}"
            if start_date and end_date else None
        ),
        "live_dates_full": (
            f"{fmt_date_full(start_date)} - {fmt_date_full(end_date)}"
            if start_date and end_date else None
        ),
        "performance_ctr": fmt_percent(ctr, 2),
        "performance_viewability": fmt_percent(viewability, 2),
        "performance_engagement_rate": fmt_percent(engagement_rate, 2),
        "performance_vcr": fmt_percent(vcr, 2),
        "planned_impressions": fmt_int(planned_impressions),
        "delivered_impressions": fmt_int(delivered_impressions),
        "total_impressions": fmt_int(total_impressions),
        "ad_value": fmt_num(ad_value, 2),
        "top_site_1_name": top_sites[0]["site"] if len(top_sites) > 0 and "site" in top_sites[0] else None,
        "top_site_1_ctr": fmt_percent(top_sites[0]["ctr"], 2) if len(top_sites) > 0 and "ctr" in top_sites[0] else None,
        "top_site_2_name": top_sites[1]["site"] if len(top_sites) > 1 and "site" in top_sites[1] else None,
        "top_site_2_ctr": fmt_percent(top_sites[1]["ctr"], 2) if len(top_sites) > 1 and "ctr" in top_sites[1] else None,
        "top_geo_1_name": top_geos[0]["geo"] if len(top_geos) > 0 and "geo" in top_geos[0] else None,
        "top_format_1_name": top_formats[0]["format"] if len(top_formats) > 0 and "format" in top_formats[0] else None,
    }

    diagnostics = {
        "metadata_found": {
            "campaign_name": meta.get("campaign_name"),
            "start_date": meta.get("start_date"),
            "end_date": meta.get("end_date"),
        },
        "candidate_tables": [asdict(c) for c in candidates[:20]],
        "selected_tables": {
            t: {
                "candidate": asdict(v["candidate"]),
                "columns": [str(col) for col in list(v["df"].columns)],
                "preview": safe_jsonable(v["df"].head(5).fillna("").to_dict(orient="records")),
            }
            for t, v in selected_tables.items()
        },
    }

    validation = validate_mapped_values(mapped)

    result = {
        "mapped_values": mapped,
        "validation": validation,
        "diagnostics": diagnostics,
    }

    return safe_jsonable(result)


PLACEHOLDER_PATTERN = re.compile(r"\{\{([^}]+)\}\}")


def replace_text_in_run(run: _Run, replacements: Dict[str, str]) -> None:
    text = run.text
    for placeholder, value in replacements.items():
        text = text.replace(f"{{{{{placeholder}}}}}", value)
    run.text = text


def replace_text_in_text_frame(text_frame, replacements: Dict[str, str]) -> None:
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            replace_text_in_run(run, replacements)

        combined = "".join(run.text for run in paragraph.runs)
        new_text = combined
        for placeholder, value in replacements.items():
            new_text = new_text.replace(f"{{{{{placeholder}}}}}", value)

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

    replacements = {}
    for placeholder, mapped_key in PLACEHOLDER_RULES.items():
        raw_value = mapped_values.get(mapped_key)
        replacements[placeholder] = "" if raw_value is None else str(raw_value)

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_in_shape(shape, replacements)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


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
        result = build_mapped_values(xls)

        clean_result = safe_jsonable(result)
        return JSONResponse(content=clean_result)

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
        result = build_mapped_values(xls)

        if not result["validation"]["is_valid"]:
            raise HTTPException(
                status_code=422,
                detail={
                    "message": "Validation failed before PPT generation",
                    "validation": result["validation"],
                    "mapped_values": result["mapped_values"],
                },
            )

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


@app.post("/process")
async def process_eoc(
    file: UploadFile = File(...),
    generate_ppt_output: bool = Form(False),
    use_server_template: bool = Form(True),
    template_file: Optional[UploadFile] = File(None),
) -> Any:
    try:
        if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            raise HTTPException(status_code=400, detail="Please upload an Excel file")

        eoc_content = await read_upload_bytes(file)
        xls = open_excel_from_bytes(eoc_content)
        result = build_mapped_values(xls)

        if not generate_ppt_output:
            clean_result = safe_jsonable(result)
            return JSONResponse(content=clean_result)

        if not result["validation"]["is_valid"]:
            raise HTTPException(
                status_code=422,
                detail={
                    "message": "Validation failed before PPT generation",
                    "validation": result["validation"],
                    "mapped_values": result["mapped_values"],
                },
            )

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
        logger.exception("process_eoc failed")
        raise HTTPException(status_code=500, detail=f"process_eoc failed: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.getenv("PORT", "8000")), reload=True)
