import base64
import json
import os
import tempfile
import traceback
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from pptx import Presentation


APP_VERSION = "12.0.0-modular-builder-poc"

app = FastAPI(
    title="PCA Modular Builder API",
    version=APP_VERSION
)

BASE_DIR = Path(__file__).resolve().parent

MODULAR_TEMPLATE_PATH = Path(
    os.getenv(
        "MODULAR_TEMPLATE_PATH",
        BASE_DIR / "templates" / "modular_sections_master.pptx"
    )
)

SECTION_REGISTRY_PATH = Path(
    os.getenv(
        "SECTION_REGISTRY_PATH",
        BASE_DIR / "section_registry" / "section_registry.json"
    )
)


class ModularOnePagerRequest(BaseModel):
    selected_sections: List[str] = Field(
        default_factory=list,
        description="Selected SECTION_ID values, for example ['TITLE_PERFORMANCE', 'CREATIVE_OVERVIEW']"
    )
    placeholder_values: Optional[Dict[str, Any]] = Field(
        default_factory=dict,
        description="Optional placeholder replacement values"
    )


def load_registry() -> Dict[str, Any]:
    if not SECTION_REGISTRY_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=f"Section registry not found at {SECTION_REGISTRY_PATH}"
        )

    with open(SECTION_REGISTRY_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_shape_text(shape) -> str:
    if not hasattr(shape, "text"):
        return ""
    return shape.text or ""


def extract_section_id_from_slide(slide) -> Optional[str]:
    for shape in slide.shapes:
        text = get_shape_text(shape).strip()
        if text.startswith("SECTION_ID:"):
            return text.replace("SECTION_ID:", "").strip()
    return None


def detect_template_sections(prs: Presentation) -> Dict[str, int]:
    detected = {}

    for index, slide in enumerate(prs.slides):
        section_id = extract_section_id_from_slide(slide)

        if section_id:
            detected[section_id] = index

    return detected


def ordered_selected_sections(
    registry: Dict[str, Any],
    selected_sections: List[str]
) -> List[str]:
    selected_clean = [s.strip().upper() for s in selected_sections if s.strip()]

    final_sections = []

    for section_id, meta in registry.items():
        if meta.get("required") is True:
            final_sections.append(section_id)

    for section_id in selected_clean:
        if section_id not in final_sections:
            final_sections.append(section_id)

    final_sections = [
        section_id for section_id in final_sections
        if section_id in registry
    ]

    final_sections.sort(
        key=lambda section_id: registry[section_id].get("default_order", 999)
    )

    return final_sections


def replace_text_in_shape(shape, placeholder_values: Dict[str, Any]) -> None:
    if not hasattr(shape, "text_frame"):
        return

    text_frame = shape.text_frame

    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if not run.text:
                continue

            new_text = run.text

            for key, value in placeholder_values.items():
                placeholder = "{{" + str(key).strip("{}") + "}}"
                replacement = "N/A" if value is None else str(value)
                new_text = new_text.replace(placeholder, replacement)

            run.text = new_text


def replace_placeholders_on_slide(slide, placeholder_values: Dict[str, Any]) -> None:
    if not placeholder_values:
        return

    for shape in slide.shapes:
        replace_text_in_shape(shape, placeholder_values)


def clone_slide(source_prs: Presentation, output_prs: Presentation, source_slide_index: int):
    source_slide = source_prs.slides[source_slide_index]
    blank_layout = output_prs.slide_layouts[6]
    new_slide = output_prs.slides.add_slide(blank_layout)

    for shape in source_slide.shapes:
        text = get_shape_text(shape).strip()

        if text.startswith("SECTION_ID:"):
            continue

        new_el = deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")

    return new_slide


def build_modular_ppt(
    selected_sections: List[str],
    placeholder_values: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    if not MODULAR_TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=f"Modular template not found at {MODULAR_TEMPLATE_PATH}"
        )

    registry = load_registry()
    source_prs = Presentation(str(MODULAR_TEMPLATE_PATH))

    output_prs = Presentation()
    output_prs.slide_width = source_prs.slide_width
    output_prs.slide_height = source_prs.slide_height

    detected_sections = detect_template_sections(source_prs)
    sections_to_build = ordered_selected_sections(registry, selected_sections)

    missing_sections = [
        section_id for section_id in sections_to_build
        if section_id not in detected_sections
    ]

    if missing_sections:
        raise HTTPException(
            status_code=400,
            detail={
                "message": "Some requested sections were not found in modular_sections_master.pptx",
                "missing_sections": missing_sections,
                "detected_sections": list(detected_sections.keys())
            }
        )

    built_sections = []

    for section_id in sections_to_build:
        source_index = detected_sections[section_id]
        new_slide = clone_slide(source_prs, output_prs, source_index)
        replace_placeholders_on_slide(new_slide, placeholder_values or {})

        built_sections.append({
            "section_id": section_id,
            "label": registry[section_id].get("label", section_id),
            "source_slide_index": source_index
        })

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        output_path = tmp.name

    output_prs.save(output_path)

    with open(output_path, "rb") as f:
        file_bytes = f.read()

    os.remove(output_path)

    return {
        "filename": "pca_modular_one_pager_v12.pptx",
        "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "file_base64": base64.b64encode(file_bytes).decode("utf-8"),
        "built_sections": built_sections
    }


@app.get("/health")
def health():
    return {
        "status": "ok",
        "app_version": APP_VERSION,
        "modular_template_exists": MODULAR_TEMPLATE_PATH.exists(),
        "section_registry_exists": SECTION_REGISTRY_PATH.exists(),
        "modular_template_path": str(MODULAR_TEMPLATE_PATH),
        "section_registry_path": str(SECTION_REGISTRY_PATH)
    }


@app.get("/list-modular-sections")
def list_modular_sections():
    registry = load_registry()

    sections = []

    for section_id, meta in registry.items():
        sections.append({
            "section_id": section_id,
            "label": meta.get("label", section_id),
            "required": meta.get("required", False),
            "default_order": meta.get("default_order", 999)
        })

    sections.sort(key=lambda item: item.get("default_order", 999))

    return {
        "app_version": APP_VERSION,
        "sections": sections
    }


@app.get("/detect-modular-template-sections")
def detect_modular_template_sections():
    if not MODULAR_TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=f"Modular template not found at {MODULAR_TEMPLATE_PATH}"
        )

    prs = Presentation(str(MODULAR_TEMPLATE_PATH))
    detected = detect_template_sections(prs)

    return {
        "app_version": APP_VERSION,
        "detected_sections": detected,
        "count": len(detected)
    }


@app.post("/generate-modular-one-pager")
def generate_modular_one_pager(request: ModularOnePagerRequest):
    try:
        return build_modular_ppt(
            selected_sections=request.selected_sections,
            placeholder_values=request.placeholder_values or {}
        )

    except HTTPException:
        raise

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail={
                "message": "Failed to generate modular one pager",
                "error": str(e)
            }
        )
