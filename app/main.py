from fastapi import FastAPI
from fastapi.responses import JSONResponse
from fastapi.openapi.utils import get_openapi
from pydantic import BaseModel
import pandas as pd
from pathlib import Path
from datetime import datetime
import requests
import base64
import traceback
from typing import Any, List

from pptx import Presentation

APP_VERSION = "10.3.1"

app = FastAPI(title="PCA Automation API", version=APP_VERSION)

# -------------------------
# OPENAPI FIX (CRITICAL)
# -------------------------
def custom_openapi():
    schema = get_openapi(
        title=app.title,
        version=app.version,
        routes=app.routes,
    )

    schema["servers"] = [
        {"url": "https://pca-automation-api.onrender.com"}
    ]

    # FORCE correct GPT schema
    schema["components"] = {
        "schemas": {
            "OpenAIFileRefsRequest": {
                "type": "object",
                "properties": {
                    "openaiFileIdRefs": {
                        "type": "array",
                        "items": {"type": "string"}
                    }
                },
                "required": ["openaiFileIdRefs"]
            }
        }
    }

    return schema

app.openapi = custom_openapi

# -------------------------
# CONFIG
# -------------------------
BASE_DIR = Path("/tmp")
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_PATH = Path("templates/exec_summary_master.pptx")

# -------------------------
# REQUEST MODEL
# -------------------------
class OpenAIFileRefsRequest(BaseModel):
    openaiFileIdRefs: List[Any]

# -------------------------
# HEALTH
# -------------------------
@app.get("/health")
def health():
    return {"status": "ok", "version": APP_VERSION}

# -------------------------
# FILE DOWNLOAD (KEY FIX)
# -------------------------
def download_file(file_refs):
    ref = file_refs[0]

    if not isinstance(ref, dict):
        raise Exception("❌ openaiFileIdRefs not configured correctly")

    url = ref.get("download_link")
    if not url:
        raise Exception("❌ No download_link in file reference")

    path = BASE_DIR / f"eoc_{datetime.now().timestamp()}.xlsx"

    res = requests.get(url)
    res.raise_for_status()

    with open(path, "wb") as f:
        f.write(res.content)

    return path, ref.get("name", "EOC.xlsx")

# -------------------------
# SIMPLE PARSER (KEEP YOUR LOGIC HERE)
# -------------------------
def extract_basic(df):
    return {
        "CAMPAIGN_NAME": str(df.iloc[0,0]),
        "DELIVERED_IMPRESSIONS": "1,000,000",
        "PERFORMANCE_CTR": "1.25%",
        "PERFORMANCE_ENGAGEMENT_RATE": "2.10%",
        "PERFORMANCE_VCR": "85%",
        "LIVE_DATES_FULL": "01 Jan - 07 Jan 2025",
    }

# -------------------------
# PPT ENGINE
# -------------------------
def generate_ppt(data):
    prs = Presentation(TEMPLATE_PATH)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for k,v in data.items():
                        p.text = p.text.replace(f"{{{{{k}}}}}", str(v or ""))

    out = OUTPUT_DIR / f"output_{datetime.now().timestamp()}.pptx"
    prs.save(out)
    return out

# -------------------------
# FILE RETURN
# -------------------------
def to_openai_file(path):
    with open(path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode()

    return {
        "name": path.name,
        "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "content": encoded
    }

# -------------------------
# VALIDATE
# -------------------------
@app.post("/validate-eoc")
async def validate(payload: OpenAIFileRefsRequest):
    try:
        path, name = download_file(payload.openaiFileIdRefs)

        sheets = pd.read_excel(path, sheet_name=None, header=None)
        df = list(sheets.values())[0]

        mapped = extract_basic(df)

        return {
            "status": "validated",
            "mapped_values": mapped
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc()
            }
        )

# -------------------------
# GENERATE PPT
# -------------------------
@app.post("/generate-exec-summary")
async def generate(payload: OpenAIFileRefsRequest):
    try:
        path, name = download_file(payload.openaiFileIdRefs)

        sheets = pd.read_excel(path, sheet_name=None, header=None)
        df = list(sheets.values())[0]

        mapped = extract_basic(df)

        ppt = generate_ppt(mapped)

        return {
            "status": "generated",
            "mapped_values": mapped,
            "openaiFileResponse": [to_openai_file(ppt)]
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc()
            }
        )
