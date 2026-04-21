from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse, FileResponse
import pandas as pd
import shutil
from pathlib import Path
from datetime import datetime
import traceback
from pptx import Presentation

app = FastAPI(title="PCA Automation API", version="8.0.0")

# -------------------------
# CONFIG
# -------------------------
BASE_DIR = Path("/tmp")
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


# -------------------------
# HEALTH
# -------------------------
@app.get("/health")
def health():
    return {"status": "ok"}


# -------------------------
# SAFE HELPERS
# -------------------------
def safe_number(df, column_name):
    try:
        val = df[column_name].dropna().iloc[0]
        return float(val)
    except:
        return None


def safe_percent(df, column_name):
    try:
        val = df[column_name].dropna().iloc[0]
        return f"{round(float(val) * 100, 2)}%"
    except:
        return None


def safe_date(df):
    try:
        return str(df.iloc[0, 1])
    except:
        return None


# -------------------------
# CORE EXTRACTION
# -------------------------
def extract_eoc_data(df: pd.DataFrame):
    result = {}

    result["CAMPAIGN_NAME"] = str(df.iloc[0, 0]) if not df.empty else ""

    result["DELIVERED_IMPRESSIONS"] = safe_number(df, "Delivered Impressions")
    result["CTR"] = safe_percent(df, "CTR")
    result["ENGAGEMENT_RATE"] = safe_percent(df, "Engagement Rate")
    result["VCR"] = safe_percent(df, "VCR")
    result["SPEND"] = safe_number(df, "Actual Spend")
    result["REPORT_DATE"] = safe_date(df)

    return result


# -------------------------
# PLACEHOLDER ENGINE
# -------------------------
def replace_text(text, data):
    if not text:
        return text

    for key, value in data.items():
        placeholder = f"{{{{{key}}}}}"
        if placeholder in text:
            text = text.replace(placeholder, str(value) if value is not None else "")
    return text


def replace_in_shape(shape, data):
    if not shape.has_text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.text = replace_text(run.text, data)


def replace_in_table(table, data):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.text = replace_text(run.text, data)


def process_slide(slide, data):
    for shape in slide.shapes:

        # Text replacement
        if shape.has_text_frame:
            replace_in_shape(shape, data)

        # Table replacement
        if shape.has_table:
            replace_in_table(shape.table, data)


def generate_ppt(template_path: Path, output_path: Path, data: dict):
    prs = Presentation(template_path)

    for slide in prs.slides:
        process_slide(slide, data)

    prs.save(output_path)


# -------------------------
# VALIDATE EOC
# -------------------------
@app.post("/validate-eoc")
async def validate_eoc(eoc_file: UploadFile = File(...)):
    try:
        path = BASE_DIR / f"eoc_{datetime.now().timestamp()}.xlsx"

        with open(path, "wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        df = pd.read_excel(path)

        mapped = extract_eoc_data(df)

        return JSONResponse(content={
            "status": "validated",
            "mapped_values": mapped
        })

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc()
            }
        )


# -------------------------
# GENERATE EXEC SUMMARY
# -------------------------
@app.post("/generate-exec-summary")
async def generate_exec_summary(
    eoc_file: UploadFile = File(...),
    template_file: UploadFile = File(...)
):
    try:
        # -------------------------
        # SAVE FILES
        # -------------------------
        eoc_path = BASE_DIR / f"eoc_{datetime.now().timestamp()}.xlsx"
        template_path = BASE_DIR / f"template_{datetime.now().timestamp()}.pptx"

        with open(eoc_path, "wb") as f:
            shutil.copyfileobj(eoc_file.file, f)

        with open(template_path, "wb") as f:
            shutil.copyfileobj(template_file.file, f)

        # -------------------------
        # EXTRACT DATA
        # -------------------------
        df = pd.read_excel(eoc_path)
        mapped_data = extract_eoc_data(df)

        # -------------------------
        # GENERATE PPT
        # -------------------------
        temp_output = OUTPUT_DIR / "temp_output.pptx"

        generate_ppt(
            template_path=template_path,
            output_path=temp_output,
            data=mapped_data
        )

        # -------------------------
        # FINAL OUTPUT (NO CACHE)
        # -------------------------
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        final = OUTPUT_DIR / f"Exec_Summary_Output_{timestamp}.pptx"

        shutil.copy2(temp_output, final)

        return FileResponse(
            path=str(final),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=final.name,
        )

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": str(e),
                "trace": traceback.format_exc()
            }
        )
