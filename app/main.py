import io
import math
import logging
from datetime import datetime
from typing import Any, Dict

import pandas as pd
import numpy as np
from numbers import Real

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse

# -----------------------
# CONFIG
# -----------------------

RULES_PATH = "PCA_GPT_Rules_Master.xlsx"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()


# -----------------------
# SAFE JSON SERIALIZER
# -----------------------

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


# -----------------------
# LOAD RULES MASTER
# -----------------------

def load_rules_master():
    try:
        xls = pd.ExcelFile(RULES_PATH, engine="openpyxl")
        df = pd.read_excel(xls, sheet_name="Placeholder Mapping")

        df.columns = [str(c).strip().lower() for c in df.columns]

        return df

    except Exception as e:
        logger.error(f"Rules Master load failed: {e}")
        return pd.DataFrame()


# -----------------------
# READ EXCEL
# -----------------------

def read_excel(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")


# -----------------------
# SIMPLE TABLE DETECTION
# -----------------------

def detect_tables(xls) -> Dict[str, pd.DataFrame]:
    tables = {}

    for sheet in xls.sheet_names:
        df = xls.parse(sheet)

        if df.empty:
            continue

        df.columns = [str(c).lower() for c in df.columns]

        # simple detection rules
        if "campaign" in df.columns:
            tables["campaign"] = df

        if "site" in df.columns:
            tables["site"] = df

        if "format" in df.columns:
            tables["format"] = df

        if "geo" in df.columns or "country" in df.columns:
            tables["geo"] = df

        if "date" in df.columns:
            tables["date"] = df

    return tables


# -----------------------
# RULE-DRIVEN MAPPING
# -----------------------

def build_mapped_values(xls):
    rules_df = load_rules_master()
    tables = detect_tables(xls)

    mapped = {}

    for _, rule in rules_df.iterrows():

        placeholder = str(rule.get("placeholder", "")).strip()
        table_type = str(rule.get("table/section", "")).lower()
        column_name = str(rule.get("source field name", "")).lower()
        logic = str(rule.get("logic", "")).lower()

        if not placeholder:
            continue

        value = None

        table = tables.get(table_type)

        if table is not None:

            df = table.copy()
            df.columns = [str(c).lower() for c in df.columns]

            # remove duplicate header rows
            df = df[df[df.columns[0]] != df.columns[0]]

            if column_name in df.columns:

                series = pd.to_numeric(df[column_name], errors="coerce").dropna()

                if not series.empty:

                    if logic == "sum":
                        value = series.sum()

                    elif logic in ["avg", "average"]:
                        value = series.mean()

                    elif "top" in logic:
                        value = series.sort_values(ascending=False).iloc[0]

                    elif "first" in logic:
                        value = series.iloc[0]

        mapped[placeholder] = value

    return {
        "mapped_values": mapped,
        "validation": {
            "is_valid": True,
            "missing_required": []
        },
        "diagnostics": {
            "tables_detected": list(tables.keys())
        }
    }


# -----------------------
# ENDPOINT
# -----------------------

@app.post("/validate-eoc")
async def validate_eoc(file: UploadFile = File(...)):
    try:
        if not file.filename.lower().endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="Upload Excel file")

        content = await file.read()
        xls = read_excel(content)

        result = build_mapped_values(xls)

        clean = safe_jsonable(result)

        return JSONResponse(content=clean)

    except Exception as e:
        logger.exception("Validation failed")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/health")
def health():
    return {"status": "ok"}
