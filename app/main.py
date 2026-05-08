import base64
import json
import os
import tempfile
import traceback
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from pptx import Presentation


APP_VERSION = "12.1.0-modular-stack-poc"

app = FastAPI(title="PCA Modular Builder API", version=APP_VERSION)

BASE_DIR = Path(__file__).resolve().parent

MODULAR_TEMPLATE_PATH = Path(
    os.getenv("MODULAR_TEMPLATE_PATH", BASE_DIR / "templates" / "modular_sections_master.pptx")
)

SECTION_REGISTRY_PATH = Path(
    os.getenv("SECTION_REGISTRY_PATH", BASE_DIR / "section_registry" / "section_registry.json")
)

EMU_PER_PX = 9525


class ModularOnePagerRequest(BaseModel):
    selected_sections: List[str] = Field(default_factory=list)
    placeholder_values: Optional[Dict[str, Any]] = Field(default_factory=dict)
    top_margin_px: int = 100
    bottom_margin_px: int = 100
    section_spacing_px: int = 60
    left_margin_px: int = 0


def px_to_emu(px: int) -> int:
    return int(px * EMU_PER_PX)


def load_registry() -> Dict[str, Any]:
    if not SECTION_REGISTRY_PATH.exists():
        raise HTTPException(status_code=500, detail=f"Section registry not found at {SECTION_REGISTRY_PATH}")

    with open(SECTION_REGISTRY_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_shape_text(shape) -> str:
    if not hasattr(shape, "text"):
        return ""
    return shape.text or ""


def is_section_id_shape(shape) -> bool:
    return get_shape_text(shape).strip().startswith("SECTION_ID:")


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


def ordered_selected_sections(registry: Dict[str, Any], selected_sections: List[str]) -> List[str]:
    selected_clean = [s.strip().upper() for s in selected_sections if s.strip()]
    final_sections = []

    for section_id, meta in registry.items():
        if meta.get("required") is True:
            final_sections.append(section_id)

    for section_id in selected_clean:
        if section_id not in final_sections:
            final_sections.append(section_id)

    final_sections = [section_id for section_id in final_sections if section_id in registry]
    final_sections.sort(key=lambda section_id: registry[section_id].get("default_order", 999))

    return final_sections


def replace_text_in_shape(shape, placeholder_values: Dict[str, Any]) -> None:
    if not hasattr(shape, "text_frame"):
        return

    for paragraph in shape.text_frame.paragraphs:
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


def get_section_bbox(slide) -> Optional[Tuple[int, int, int, int]]:
    shapes = [shape for shape in slide.shapes if not is_section_id_shape(shape)]

    if not shapes:
        return None

    min_left = min(shape.left for shape in shapes)
    min_top = min(shape.top for shape in shapes)
    max_right = max(shape.left + shape.width for shape in shapes)
    max_bottom = max(shape.top + shape.height for shape in shapes)

    return min_left, min_top, max_right, max_bottom


def copy_shape_to_slide(source_shape, target_slide, new_left: int, new_top: int):
    new_el = deepcopy(source_shape.element)
    target_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")

    copied_shape = list(target_slide.shapes)[-1]

    try:
        copied_shape.left = new_left
        copied_shape.top = new_top
    except Exception:
        pass

    return copied_shape


def build_stacked_modular_ppt(
    selected_sections: List[str],
    placeholder_values: Optional[Dict[str, Any]] = None,
    top_margin_px: int = 100,
    bottom_margin_px: int = 100,
    section_spacing_px: int = 60,
    left_margin_px: int = 0
) -> Dict[str, Any]:

    if not MODULAR_TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail=f"Modular template not found at {MODULAR_TEMPLATE_PATH}")

    registry = load_registry()
    source_prs = Presentation(str(MODULAR_TEMPLATE_PATH))
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

    section_data = []

    for section_id in sections_to_build:
        source_slide = source_prs.slides[detected_sections[section_id]]
        bbox = get_section_bbox(source_slide)

        if not bbox:
            continue

        min_left, min_top, max_right, max_bottom = bbox
        section_height = max_bottom - min_top

        section_data.append({
            "section_id": section_id,
            "label": registry[section_id].get("label", section_id),
            "slide": source_slide,
            "bbox": bbox,
            "height": section_height
        })

    top_margin = px_to_emu(top_margin_px)
    bottom_margin = px_to_emu(bottom_margin_px)
    spacing = px_to_emu(section_spacing_px)
    left_margin = px_to_emu(left_margin_px)

    total_height = top_margin + bottom_margin

    for i, section in enumerate(section_data):
        total_height += section["height"]
        if i < len(section_data) - 1:
            total_height += spacing

    output_prs = Presentation()
    output_prs.slide_width = source_prs.slide_width
    output_prs.slide_height = total_height

    blank_layout = output_prs.slide_layouts[6]
    output_slide = output_prs.slides.add_slide(blank_layout)

    cursor_y = top_margin
    built_sections = []

    for section in section_data:
        source_slide = section["slide"]
        min_left, min_top, max_right, max_bottom = section["bbox"]

        for shape in source_slide.shapes:
            if is_section_id_shape(shape):
                continue

            new_left = left_margin + (shape.left - min_left)
            new_top = cursor_y + (shape.top - min_top)

            copy_shape_to_slide(shape, output_slide, new_left, new_top)

        built_sections.append({
            "section_id": section["section_id"],
            "label": section["label"],
            "height_emu": section["height"]
        })

        cursor_y += section["height"] + spacing

    replace_placeholders_on_slide(output_slide, placeholder_values or {})

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        output_path = tmp.name

    output_prs.save(output_path)

    with open(output_path, "rb") as f:
        file_bytes = f.read()

    os.remove(output_path)

    return {
        "filename": "pca_modular_stacked_one_pager_v12.pptx",
        "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "file_base64": base64.b64encode(file_bytes).decode("utf-8"),
        "built_sections": built_sections,
        "slide_height_emu": total_height,
        "slide_height_px_approx": round(total_height / EMU_PER_PX)
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
        raise HTTPException(status_code=500, detail=f"Modular template not found at {MODULAR_TEMPLATE_PATH}")

    prs = Presentation(str(MODULAR_TEMPLATE_PATH))
    detected = detect_template_sections(prs)

    return {
        "app_version": APP_VERSION,
        "detected_sections": detected,
        "count": len(detected)
    }


@app.post("/generate-modular-one-pager-stacked")
def generate_modular_one_pager_stacked(request: ModularOnePagerRequest):
    try:
        return build_stacked_modular_ppt(
            selected_sections=request.selected_sections,
            placeholder_values=request.placeholder_values or {},
            top_margin_px=request.top_margin_px,
            bottom_margin_px=request.bottom_margin_px,
            section_spacing_px=request.section_spacing_px,
            left_margin_px=request.left_margin_px
        )

    except HTTPException:
        raise

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail={
                "message": "Failed to generate stacked modular one pager",
                "error": str(e)
            }
        )
