from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse, FileResponse
import pandas as pd
import shutil
from pathlib import Path
from datetime import datetime
import traceback

app = FastAPI(title="PCA Automation API", version="7.0.0")

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
# CORE: EOC PARSER
# -------------------------
def extract_eoc_data(df: pd.DataFrame):
    """
    Core extraction logic from EOC.
    Expand this as needed.
    """

    result = {}

    try:
        # Campaign Name
        result["campaign_name"] = str(df.iloc[0, 0])

        # Basic metrics (safe access)
        result["impressions"] = safe_number(df, "Delivered Impressions")
        result["ctr"] = safe_percent(df, "CTR")
        result["engagement_rate"] = safe_percent(df, "Engagement Rate")
        result["vcr"] = safe_percent(df, "VCR")

        result["spend"] = safe_number(df, "Actual Spend")

        # Dates
        result["report_date"] = safe_date(df)

    except Exception as e:
        print("Extraction warning:", e)

    return result


def safe_number(df, column_name):
    try:
        val = df[column_name].dropna().iloc[0]
        return float(val)
    except:
        return None


def safe_percent(df, column_name):
    try:
        val = df[column_name].dropna().iloc[0]
        return round(float(val) * 100, 2)
    except:
        return None


def safe_date(df):
    try:
        val = df.iloc[0, 1]
        return str(val)
    except:
        return None


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
# PPT GENERATION ENGINE
# -------------------------
def generate_ppt_from_template(template_path: Path, output_path: Path, data: dict):
    """
    Placeholder logic.
    Replace with python-pptx mapping later.
    """

    # TEMP: just copy template
    shutil.copy2(template_path, output_path)


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
        # PROCESS DATA
        # -------------------------
        df = pd.read_excel(eoc_path)
        mapped_data = extract_eoc_data(df)

        # -------------------------
        # GENERATE PPT
        # -------------------------
        temp_output = OUTPUT_DIR / "temp_output.pptx"

        generate_ppt_from_template(
            template_path=template_path,
            output_path=temp_output,
            data=mapped_data
        )

        # -------------------------
        # FINAL OUTPUT (NO CACHE BUG)
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
