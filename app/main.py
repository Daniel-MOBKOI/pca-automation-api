from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
import shutil
import os
from pathlib import Path
from pptx import Presentation
import openpyxl
import tempfile

app = FastAPI()

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# -------------------------
# Health Check
# -------------------------
@app.get("/health")
def health():
    return {"status": "ok"}


# -------------------------
# SIMPLE TEST PARSER (V1)
# -------------------------
def extract_basic_data(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    data = {
        "campaign_name": "N/A",
        "impressions": "N/A",
        "ctr": "N/A",
        "vcr": "N/A",
        "spend": "N/A"
    }

    for row in ws.iter_rows(values_only=True):
        row_str = [str(x).lower() if x else "" for x in row]

        if "campaign" in row_str:
            data["campaign_name"] = str(row[1])

        if "impressions" in row_str:
            data["impressions"] = str(row[1])

        if "ctr" in row_str:
            data["ctr"] = str(row[1])

        if "vcr" in row_str:
            data["vcr"] = str(row[1])

        if "spend" in row_str:
            data["spend"] = str(row[1])

    return data


# -------------------------
# PPT REPLACER
# -------------------------
def replace_text(ppt_path, output_path, data):
    prs = Presentation(ppt_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text

                        text = text.replace("{{CAMPAIGN_NAME}}", data["campaign_name"])
                        text = text.replace("{{DELIVERED_IMPRESSIONS}}", data["impressions"])
                        text = text.replace("{{PERFORMANCE_CTR}}", data["ctr"])
                        text = text.replace("{{PERFORMANCE_VCR}}", data["vcr"])
                        text = text.replace("{{CAMPAIGN_BUDGET}}", data["spend"])

                        run.text = text

    prs.save(output_path)


# -------------------------
# VALIDATE ENDPOINT
# -------------------------
@app.post("/validate-eoc")
async def validate_eoc(eoc_file: UploadFile = File(...)):
    with tempfile.TemporaryDirectory() as tmp:
        file_path = os.path.join(tmp, eoc_file.filename)

        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(eoc_file.file, buffer)

        data = extract_basic_data(file_path)

        return JSONResponse(content=data)


# -------------------------
# GENERATE PPT
# -------------------------
@app.post("/generate-exec-summary")
async def generate_exec_summary(
    eoc_file: UploadFile = File(...),
    template_file: UploadFile = File(...)
):
    with tempfile.TemporaryDirectory() as tmp:

        eoc_path = os.path.join(tmp, eoc_file.filename)
        template_path = os.path.join(tmp, template_file.filename)

        with open(eoc_path, "wb") as buffer:
            shutil.copyfileobj(eoc_file.file, buffer)

        with open(template_path, "wb") as buffer:
            shutil.copyfileobj(template_file.file, buffer)

        data = extract_basic_data(eoc_path)

        output_file = os.path.join(tmp, "output.pptx")

        replace_text(template_path, output_file, data)

        final_path = OUTPUT_DIR / "Exec_Summary_Output.pptx"
        shutil.copy(output_file, final_path)

        return FileResponse(
            final_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename="Exec_Summary.pptx"
        )
