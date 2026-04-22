from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
import shutil
from pathlib import Path
from datetime import datetime
from pptx import Presentation
import openpyxl
import traceback

app = FastAPI()

# -------------------------
# PATHS
# -------------------------
BASE_DIR = Path("/tmp")
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# ✅ TEMPLATE (FROM REPO ROOT)
TEMPLATE_PATH = Path("Executive Summary_PCA_One Pager_MASTER.pptx")


# -------------------------
# HELPERS
# -------------------------
def safe_number(val):
    try:
        return float(val)
    except:
        return None


def clean_string(val):
    if val is None:
        return None
    return str(val).strip()


# -------------------------
# CORE EXTRACTION (STABLE)
# -------------------------
def extract_eoc_data(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)

    result = {
        "CAMPAIGN_NAME": None,
        "DELIVERED_IMPRESSIONS": None,
        "PERFORMANCE_CTR": None,
        "PERFORMANCE_ENGAGEMENT_RATE": None,
        "PERFORMANCE_VCR": None,
        "CAMPAIGN_BUDGET": None,
        "CAMPAIGN_MARKETS": None,
        "LIVE_DATES_FULL": None,
    }

    try:
        for sheet in wb.worksheets:

            for row in sheet.iter_rows(values_only=True):

                row_values = [str(x).lower() if x else "" for x in row]

                # Campaign Name
                if "campaign" in row_values and not result["CAMPAIGN_NAME"]:
                    result["CAMPAIGN_NAME"] = clean_string(row[1])

                # Impressions
                if "impressions" in row_values:
                    result["DELIVERED_IMPRESSIONS"] = safe_number(row[1])

                # CTR
                if "ctr" in row_values:
                    result["PERFORMANCE_CTR"] = clean_string(row[1])

                # Engagement Rate
                if "engagement" in row_values:
                    result["PERFORMANCE_ENGAGEMENT_RATE"] = clean_string(row[1])

                # VCR
                if "vcr" in row_values:
                    result["PERFORMANCE_VCR"] = clean_string(row[1])

                # Spend / Budget
                if "spend" in row_values or "budget" in row_values:
                    result["CAMPAIGN_BUDGET"] = safe_number(row[1])

                # Markets
                if "market" in row_values or "geo" in row_values:
                    result["CAMPAIGN_MARKETS"] = clean_string(row[1])

                # Dates
                if "date" in row_values:
                    result["LIVE_DATES_FULL"] = clean_string(row[1])

    except Exception as e:
        print("EXTRACTION ERROR:", e)

    return result


# -------------------------
# PPT POPULATION
# -------------------------
def populate_ppt(data, output_path):
    prs = Presentation(TEMPLATE_PATH)

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            for paragraph in shape.text_frame.paragraphs:

                text = paragraph.text

                for key, value in data.items():
                    placeholder = f"{{{{{key}}}}}"

                    if placeholder in text:
                        replacement = str(value) if value is not None else "N/A"
                        text = text.replace(placeholder, replacement)

                paragraph.text = text

    prs.save(output_path)


# -------------------------
# ENDPOINTS
# -------------------------
@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/generate")
async def generate(file: UploadFile = File(...)):
    try:
        file_path = BASE_DIR / file.filename

        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        data = extract_eoc_data(file_path)

        output_file = OUTPUT_DIR / f"output_{datetime.now().timestamp()}.pptx"

        populate_ppt(data, output_file)

        return FileResponse(
            output_file,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename="Exec_Summary.pptx",
        )

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e), "trace": traceback.format_exc()},
        )
