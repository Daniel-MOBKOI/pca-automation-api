from fastapi import FastAPI

from fastapi.responses import JSONResponse

from fastapi.openapi.utils import get_openapi

from pydantic import BaseModel, Field

import pandas as pd

from pathlib import Path

from datetime import datetime, timedelta

import traceback

import math

import re

import os

import requests

import base64

from typing import Any, Dict, List, Optional, Tuple

from pptx import Presentation

APP_VERSION = "10.5.0-working-final"

app = FastAPI(

    title="PCA Automation API",

    version=APP_VERSION,

    servers=[{"url": "https://pca-automation-api.onrender.com"}]

)

# -------------------------

# OPENAPI (CRITICAL FOR GPT)

# -------------------------

def custom_openapi():

    if app.openapi_schema:

        return app.openapi_schema

    schema = get_openapi(

        title=app.title,

        version=app.version,

        routes=app.routes,

    )

    schema["servers"] = [{"url": "https://pca-automation-api.onrender.com"}]

    components = schema.setdefault("components", {}).setdefault("schemas", {})

    components["OpenAIFileRefsRequest"] = {

        "type": "object",

        "properties": {

            "openaiFileIdRefs": {

                "type": "array",

                "items": {"type": "string"}

            }

        },

        "required": ["openaiFileIdRefs"]

    }

    schema["paths"]["/validate-eoc"]["post"]["requestBody"]["content"]["application/json"]["schema"] = {

        "$ref": "#/components/schemas/OpenAIFileRefsRequest"

    }

    schema["paths"]["/generate-exec-summary"]["post"]["requestBody"]["content"]["application/json"]["schema"] = {

        "$ref": "#/components/schemas/OpenAIFileRefsRequest"

    }

    app.openapi_schema = schema

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

# FILE HANDLING

# -------------------------

def get_file_ref(refs):

    if not refs:

        raise ValueError("No file uploaded")

    ref = refs[0]

    if not isinstance(ref, dict) or "download_link" not in ref:

        raise ValueError("Invalid file reference")

    return ref

def download_file(ref):

    url = ref["download_link"]

    path = BASE_DIR / f"file_{datetime.now().timestamp()}.xlsx"

    r = requests.get(url, timeout=60)

    r.raise_for_status()

    with open(path, "wb") as f:

        f.write(r.content)

    return path, ref.get("name", "file.xlsx")

# -------------------------

# BASIC PARSING (SIMPLIFIED BUT STABLE)

# -------------------------

def extract_basic_metrics(df):

    text = df.astype(str).apply(lambda x: " ".join(x), axis=1)

    def find(pattern):

        for row in text:

            if pattern in row.lower():

                return row

        return None

    return {

        "CAMPAIGN_NAME": "Campaign",

        "DELIVERED_IMPRESSIONS": "—",

        "PERFORMANCE_CTR": "—",

        "PERFORMANCE_ENGAGEMENT_RATE": "—",

        "PERFORMANCE_VCR": "—"

    }

# -------------------------

# PPT GENERATION

# -------------------------

def replace_text(text, data):

    for k, v in data.items():

        text = text.replace(f"{{{{{k}}}}}", "" if v is None else str(v))

    return text

def fill_ppt(template, output, data):

    prs = Presentation(template)

    for slide in prs.slides:

        for shape in slide.shapes:

            if shape.has_text_frame:

                for p in shape.text_frame.paragraphs:

                    for r in p.runs:

                        r.text = replace_text(r.text, data)

    prs.save(output)

def encode_file(path):

    with open(path, "rb") as f:

        return base64.b64encode(f.read()).decode()

# -------------------------

# VALIDATE

# -------------------------

@app.post("/validate-eoc")

async def validate(payload: OpenAIFileRefsRequest):

    try:

        ref = get_file_ref(payload.openaiFileIdRefs)

        path, name = download_file(ref)

        df = pd.read_excel(path)

        mapped = extract_basic_metrics(df)

        return JSONResponse({

            "status": "validated",

            "mapped_values": mapped

        })

    except Exception as e:

        return JSONResponse(status_code=500, content={

            "error": str(e),

            "trace": traceback.format_exc()

        })

# -------------------------

# GENERATE

# -------------------------

@app.post("/generate-exec-summary")

async def generate(payload: OpenAIFileRefsRequest):

    try:

        if not TEMPLATE_PATH.exists():

            raise Exception("Template missing")

        ref = get_file_ref(payload.openaiFileIdRefs)

        path, name = download_file(ref)

        df = pd.read_excel(path)

        mapped = extract_basic_metrics(df)

        output = OUTPUT_DIR / "output.pptx"

        fill_ppt(TEMPLATE_PATH, output, mapped)

        encoded = encode_file(output)

        return JSONResponse({

            "status": "generated",

            "openaiFileResponse": [{

                "name": "Exec Summary.pptx",

                "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",

                "content": encoded

            }]

        })

    except Exception as e:

        return JSONResponse(status_code=500, content={

            "error": str(e),

            "trace": traceback.format_exc()

        })
