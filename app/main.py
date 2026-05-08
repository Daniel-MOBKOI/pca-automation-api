import base64
import json
import os
import tempfile
import traceback
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from pptx import Presentation
from starlette.background import BackgroundTask


APP_VERSION = "12.5.0-section-alignment-poc"

app = FastAPI(title="PCA Modular Builder API", version=APP_VERSION)

BASE_DIR = Path(__file__).resolve().parent

MODULAR_TEMPLATE_PATH = Path(
    os.getenv("MODULAR_TEMPLATE_PATH", BASE_DIR / "templates" / "modular_sections_master.pptx")
)

SECTION_REGISTRY_PATH = Path(
    os.getenv("SECTION_REGISTRY_PATH", BASE_DIR / "section_registry" / "section_registry.json")
)

EMU_PER_PX = 9525
PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class ModularOnePagerRequest(BaseModel):
    selected_sections: List[str] = Field(default_factory=list)
    placeholder_values: Optional[Dict[str, Any]] = Field(default_factory=dict)
    top_margin_px: int = 100
    bottom_margin_px: int = 100
    section_spacing_px: int = 60


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


def find_main_group_shape(slide):
    candidates = []

    for shape in slide.shapes:
        if is_section_id_shape(shape):
            continue

        is_group = shape.shape_type == 6
        area = int(shape.width) * int(shape.height)

        candidates.append({
            "shape": shape,
            "is_group": is_group,
            "area": area
        })

    if not candidates:
        return None

    group_candidates = [c for c in candidates if c["is_group"]]

    if group_candidates:
        return max(group_candidates, key=lambda c: c["area"])["shape"]

    return max(candidates, key=lambda c: c["area"])["shape"]


def replace_text_in_shape(shape, placeholder_values: Dict[str, Any]) -> None:
    if hasattr(shape, "text_frame"):
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

    if hasattr(shape, "shapes"):
        for subshape in shape.shapes:
            replace_text_in_shape(subshape, placeholder_values)


def replace_placeholders_on_slide(slide, placeholder_values: Dict[str, Any]) -> None:
    if not placeholder_values:
        return

    for shape in slide.shapes:
        replace_text_in_shape(shape, placeholder_values)


def create_output_presentation_with_source_theme(source_template_path: Path, slide_height: int) -> Presentation:
    output_prs = Presentation(str(source_template_path))
    output_prs.slide_height = slide_height

    slide_id_list = output_prs.slides._sldIdLst

    for slide_id in list(slide_id_list):
        output_prs.part.drop_rel(slide_id.rId)
        slide_id_list.remove(slide_id)

    return output_prs


def collect_relationship_ids_from_element(element) -> List[str]:
    rel_ids = set()

    for node in element.iter():
        for attr_name, attr_value in node.attrib.items():
            if attr_name.startswith("{" + REL_NS + "}"):
                rel_ids.add(attr_value)

    return list(rel_ids)


def remap_relationship_ids_in_element(element, rel_id_map: Dict[str, str]) -> None:
    for node in element.iter():
        for attr_name, attr_value in list(node.attrib.items()):
            if attr_value in rel_id_map:
                node.attrib[attr_name] = rel_id_map[attr_value]


def copy_relationships_for_element(source_slide, target_slide, copied_element) -> Dict[str, str]:
    rel_id_map = {}
    rel_ids = collect_relationship_ids_from_element(copied_element)

    for old_rid in rel_ids:
        try:
            source_rel = source_slide.part.rels[old_rid]
        except KeyError:
            continue

        try:
            if getattr(source_rel, "is_external", False):
                new_rid = target_slide.part.relate_to(
                    source_rel.target_ref,
                    source_rel.reltype,
                    is_external=True
                )
            else:
                new_rid = target_slide.part.relate_to(
                    source_rel.target_part,
                    source_rel.reltype
                )

            rel_id_map[old_rid] = new_rid

        except Exception:
            continue

    remap_relationship_ids_in_element(copied_element, rel_id_map)

    return rel_id_map


def copy_group_to_slide_relationship_safe(source_shape, source_slide, target_slide, new_left: int, new_top: int):
    copied_element = deepcopy(source_shape.element)

    copy_relationships_for_element(
        source_slide=source_slide,
        target_slide=target_slide,
        copied_element=copied_element
    )

    target_slide.shapes._spTree.insert_element_before(copied_element, "p:extLst")

    copied_shape = list(target_slide.shapes)[-1]

    try:
        copied_shape.left = new_left
        copied_shape.top = new_top
    except Exception:
        pass

    return copied_shape


def calculate_section_left(section_id: str, section_width: int, slide_width: int) -> int:
    if section_id == "TITLE_OVERVIEW":
        return int(slide_width - section_width)

    return int((slide_width - section_width) / 2)


def build_grouped_stacked_modular_ppt(
    selected_sections: List[str],
    placeholder_values: Optional[Dict[str, Any]] = None,
    top_margin_px: int = 100,
    bottom_margin_px: int = 100,
    section_spacing_px: int = 60
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
        group_shape = find_main_group_shape(source_slide)

        if group_shape is None:
            continue

        section_data.append({
            "section_id": section_id,
            "label": registry[section_id].get("label", section_id),
            "slide": source_slide,
            "shape": group_shape,
            "width": group_shape.width,
            "height": group_shape.height
        })

    top_margin = px_to_emu(top_margin_px)
    bottom_margin = px_to_emu(bottom_margin_px)
    spacing = px_to_emu(section_spacing_px)

    total_height = top_margin + bottom_margin

    for i, section in enumerate(section_data):
        total_height += section["height"]
        if i < len(section_data) - 1:
            total_height += spacing

    output_prs = create_output_presentation_with_source_theme(
        source_template_path=MODULAR_TEMPLATE_PATH,
        slide_height=total_height
    )

    slide_width = output_prs.slide_width

    blank_layout = output_prs.slide_layouts[6]
    output_slide = output_prs.slides.add_slide(blank_layout)

    cursor_y = top_margin
    built_sections = []

    for section in section_data:
        section_id = section["section_id"]

        section_left = calculate_section_left(
            section_id=section_id,
            section_width=section["width"],
            slide_width=slide_width
        )

        copy_group_to_slide_relationship_safe(
            source_shape=section["shape"],
            source_slide=section["slide"],
            target_slide=output_slide,
            new_left=section_left,
            new_top=cursor_y
        )

        built_sections.append({
            "section_id": section_id,
            "label": section["label"],
            "alignment": "right" if section_id == "TITLE_OVERVIEW" else "center",
            "left_px_approx": round(section_left / EMU_PER_PX),
            "width_px_approx": round(section["width"] / EMU_PER_PX),
            "height_px_approx": round(section["height"] / EMU_PER_PX)
        })

        cursor_y += section["height"] + spacing

    replace_placeholders_on_slide(output_slide, placeholder_values or {})

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        output_path = tmp.name

    output_prs.save(output_path)

    with open(output_path, "rb") as f:
        file_bytes = f.read()

    return {
        "output_path": output_path,
        "filename": "pca_modular_grouped_stacked_one_pager_v12.pptx",
        "mime_type": PPTX_MIME_TYPE,
        "file_base64": base64.b64encode(file_bytes).decode("utf-8"),
        "built_sections": built_sections,
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


@app.get("/inspect-modular-groups")
def inspect_modular_groups():
    if not MODULAR_TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail=f"Modular template not found at {MODULAR_TEMPLATE_PATH}")

    prs = Presentation(str(MODULAR_TEMPLATE_PATH))
    detected = detect_template_sections(prs)

    results = []

    for section_id, slide_index in detected.items():
        slide = prs.slides[slide_index]
        main_group = find_main_group_shape(slide)

        results.append({
            "section_id": section_id,
            "slide_index": slide_index,
            "main_shape_found": main_group is not None,
            "main_shape_type": str(main_group.shape_type) if main_group else None,
            "main_shape_width_px_approx": round(main_group.width / EMU_PER_PX) if main_group else None,
            "main_shape_height_px_approx": round(main_group.height / EMU_PER_PX) if main_group else None
        })

    return {
        "app_version": APP_VERSION,
        "sections": results
    }


@app.post("/generate-modular-one-pager-grouped-stacked")
def generate_modular_one_pager_grouped_stacked(request: ModularOnePagerRequest):
    try:
        result = build_grouped_stacked_modular_ppt(
            selected_sections=request.selected_sections,
            placeholder_values=request.placeholder_values or {},
            top_margin_px=request.top_margin_px,
            bottom_margin_px=request.bottom_margin_px,
            section_spacing_px=request.section_spacing_px
        )

        output_path = result.pop("output_path")

        if os.path.exists(output_path):
            os.remove(output_path)

        return result

    except HTTPException:
        raise

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail={
                "message": "Failed to generate grouped stacked modular one pager",
                "error": str(e)
            }
        )


@app.post("/download-modular-one-pager-grouped-stacked")
def download_modular_one_pager_grouped_stacked(request: ModularOnePagerRequest):
    try:
        result = build_grouped_stacked_modular_ppt(
            selected_sections=request.selected_sections,
            placeholder_values=request.placeholder_values or {},
            top_margin_px=request.top_margin_px,
            bottom_margin_px=request.bottom_margin_px,
            section_spacing_px=request.section_spacing_px
        )

        output_path = result["output_path"]
        filename = result["filename"]

        return FileResponse(
            path=output_path,
            filename=filename,
            media_type=PPTX_MIME_TYPE,
            background=BackgroundTask(
                lambda: os.remove(output_path) if os.path.exists(output_path) else None
            )
        )

    except HTTPException:
        raise

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail={
                "message": "Failed to download grouped stacked modular one pager",
                "error": str(e)
            }
        )
