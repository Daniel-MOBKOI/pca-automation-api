from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
import shutil
from pathlib import Path
from datetime import datetime
from pptx import Presentation
import openpyxl
import traceback

app = FastAPI()

BASE_DIR = Path("/tmp")
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_PATH = BASE_DIR / "template.pptx"


# -------------------------
# HELPERS
# -------------------------
def find_value(ws, keyword):
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell and keyword.lower() in str(cell).lower():
                return cell
    return None


def safe_number(val):
    try:
        return float(val)
    except:
        return None


# -------------------------
# CORE EXTRACTION (WORKING VERSION)
# -------------------------
def extract_eoc_data(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)

    result = {
        "CAMPAIGN_NAME": None,
        "DELIVERED_IMPRESSIONS": None,
        "CTR": None,
        "ENGAGEMENT_RATE": None,
        "VCR": None,
        "SPEND": None,
        "MARKETS": None,
        "LIVE_DATES": None,
    }

    try:
        for sheet in wb.worksheets:

            for row in sheet.iter_rows(values_only=True):

                row_values = [str(x).lower() if x else "" for x in row]

                # Campaign Name
                if "campaign" in row_values and not result["CAMPAIGN_NAME"]:
                    result["CAMPAIGN_NAME"] = row[1]

                # Impressions
                if "impressions" in row_values:
                    result["DELIVERED_IMPRESSIONS"] = safe_number(row[1])

                # CTR
                if "ctr" in row_values:
                    result["CTR"] = row[1]

                # Engagement
                if "engagement" in row_values:
                    result["ENGAGEMENT_RATE"] = row[1]

                # VCR
                if "vcr" in row_values:
                    result["VCR"] = row[1]

                # Spend
                if "spend" in row_values or "budget" in row_values:
                    result["SPEND"] = safe_number(row[1])

                # Markets
                if "market" in row_values or "geo" in row_values:
                    result["MARKETS"] = row[1]

                # Dates
                if "date" in row_values:
                    result["LIVE_DATES"] = row[1]

    except Exception as e:
        print("ERROR:", e)

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
                for key, value in data.items():
                    if value is None:
                        value = "N/A"
                    if f"{{{{{key}}}}}" in paragraph.text:
                        paragraph.text = paragraph.text.replace(f"{{{{{key}}}}}", str(value))

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
