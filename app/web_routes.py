"""
web_routes.py
==============
Web-app endpoints for the PCA Automation Generator.

Lives at: app/web_routes.py (alongside app/main.py)

v6 changes:
- Claude API exec summary: /api/validate now calls Anthropic claude-haiku-4-5
  to write a human-narrative exec summary from validated campaign data.
  Result is capped at 550 characters and returned as `claude_exec_summary`.
  Falls back silently to None if key is missing or call fails.
- Batch upload: /api/generate is unchanged — frontend calls it once per EOC
  in parallel. No backend changes needed for batch.
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sqlite3
import tempfile
import traceback
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Optional

import httpx
from authlib.integrations.starlette_client import OAuth, OAuthError
from fastapi import (
    APIRouter,
    Depends,
    FastAPI,
    File,
    Form,
    HTTPException,
    Request,
    UploadFile,
)
from fastapi.responses import JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware


# ----- Paths ----------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent
WEB_TEMPLATES_DIR = APP_DIR / "web_templates"
STATIC_DIR = APP_DIR / "static"
STATIC_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = Path(os.getenv("PCA_DB_PATH", APP_DIR / "pca.db"))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

GENERATED_FILES_DIR = Path(
    os.getenv("GENERATED_FILES_DIR", tempfile.gettempdir())
)
GENERATED_FILES_DIR.mkdir(parents=True, exist_ok=True)

FILE_RETENTION_DAYS = int(os.getenv("FILE_RETENTION_DAYS", "30"))
CLEANUP_INTERVAL_SECONDS = 24 * 60 * 60

PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
EXEC_SUMMARY_MAX_CHARS = 550


# ----- Database -------------------------------------------------------------


def db_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def db_init() -> None:
    with db_connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                email TEXT PRIMARY KEY,
                name TEXT,
                picture TEXT,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS runs (
                run_id TEXT PRIMARY KEY,
                user_email TEXT NOT NULL,
                created_at TEXT NOT NULL,
                eoc_filename TEXT,
                output_type TEXT,
                selected_sections TEXT,
                deck_mode TEXT,
                exec_summary TEXT,
                one_pager_filename TEXT,
                deck_filename TEXT,
                summary_json TEXT,
                status TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_runs_user
                ON runs(user_email, created_at DESC);
            """
        )


db_init()


# ----- File cleanup ---------------------------------------------------------


def cleanup_old_files() -> Dict[str, int]:
    cutoff_iso = (datetime.utcnow() - timedelta(days=FILE_RETENTION_DAYS)).isoformat()
    deleted_files = 0
    missing_files = 0
    errors = 0

    try:
        with db_connect() as conn:
            old_runs = conn.execute(
                """SELECT run_id, one_pager_filename, deck_filename
                   FROM runs
                   WHERE created_at < ?
                     AND (one_pager_filename IS NOT NULL OR deck_filename IS NOT NULL)""",
                (cutoff_iso,),
            ).fetchall()

        for row in old_runs:
            for column in ("one_pager_filename", "deck_filename"):
                filename = row[column]
                if not filename:
                    continue
                file_path = GENERATED_FILES_DIR / filename
                try:
                    if file_path.exists():
                        file_path.unlink()
                        deleted_files += 1
                    else:
                        missing_files += 1
                except Exception:
                    errors += 1
                    traceback.print_exc()

            try:
                with db_connect() as conn:
                    conn.execute(
                        """UPDATE runs
                           SET one_pager_filename = NULL, deck_filename = NULL
                           WHERE run_id = ?""",
                        (row["run_id"],),
                    )
            except Exception:
                errors += 1
                traceback.print_exc()

    except Exception:
        traceback.print_exc()
        errors += 1

    if deleted_files or missing_files or errors:
        print(
            f"[cleanup] deleted={deleted_files} "
            f"already_missing={missing_files} errors={errors} "
            f"(retention={FILE_RETENTION_DAYS}d)"
        )
    return {"deleted": deleted_files, "missing": missing_files, "errors": errors}


async def _cleanup_loop() -> None:
    while True:
        try:
            cleanup_old_files()
        except Exception:
            traceback.print_exc()
        await asyncio.sleep(CLEANUP_INTERVAL_SECONDS)


_cleanup_task_started = False


def _ensure_cleanup_task_running() -> None:
    global _cleanup_task_started
    if _cleanup_task_started:
        return
    _cleanup_task_started = True
    try:
        loop = asyncio.get_event_loop()
        loop.create_task(_cleanup_loop())
        print(f"[cleanup] background task scheduled (retention={FILE_RETENTION_DAYS}d)")
    except Exception:
        traceback.print_exc()
        cleanup_old_files()


# ----- Smart filename helpers -----------------------------------------------


def _safe_filename_part(s: str, max_len: int = 60) -> str:
    """
    Slugify a string for use in a filename:
      - replace any run of non-alphanumeric chars with "_"
      - collapse repeated underscores
      - trim leading/trailing underscores
      - cap length
    """
    if not s:
        return "PCA"
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    cleaned = re.sub(r"_+", "_", cleaned)
    return cleaned[:max_len] or "PCA"


def _smart_filename(campaign_name: str, doc_type: str) -> str:
    """
    Build the desired download filename:
        <CampaignName>_<Type>_<DDMMYY>.pptx
    e.g. Prada_Eyewear_FW25_One_Pager_290526.pptx
    """
    date_part = datetime.utcnow().strftime("%d%m%y")
    name = _safe_filename_part(campaign_name or "PCA")
    return f"{name}_{doc_type}_{date_part}.pptx"


def _rename_generated_file(original_filename: str, new_filename: str) -> str:
    """
    Rename a freshly-generated file in GENERATED_FILES_DIR.
    Falls back to the original filename if the rename fails.
    """
    try:
        src = GENERATED_FILES_DIR / Path(original_filename).name
        if not src.exists():
            return Path(original_filename).name

        dst = GENERATED_FILES_DIR / new_filename
        if dst.exists() and src != dst:
            stem = dst.stem
            counter = 2
            while True:
                candidate = GENERATED_FILES_DIR / f"{stem}_{counter}.pptx"
                if not candidate.exists():
                    dst = candidate
                    break
                counter += 1

        src.rename(dst)
        return dst.name
    except Exception:
        traceback.print_exc()
        return Path(original_filename).name


# ----- Claude API exec summary ----------------------------------------------


def _truncate_to_sentence(text: str, max_chars: int) -> str:
    """
    Truncate text to at most max_chars characters, preferring to cut at a
    sentence boundary so the result still reads naturally.
    """
    if len(text) <= max_chars:
        return text
    # Try to cut at the last sentence-ending punctuation within the limit
    chunk = text[:max_chars]
    for punct in ('. ', '! ', '? '):
        pos = chunk.rfind(punct)
        if pos > max_chars // 2:
            return chunk[:pos + 1].rstrip()
    # No clean sentence boundary — hard truncate at last space
    pos = chunk.rfind(' ')
    return (chunk[:pos] if pos > 0 else chunk).rstrip()


async def _claude_exec_summary(summary_inputs: Dict[str, Any]) -> Optional[str]:
    """
    Call Claude claude-haiku-4-5 to write a natural exec summary paragraph.
    Returns None on any failure so the frontend falls back to templates.
    Result is capped at EXEC_SUMMARY_MAX_CHARS (550) characters.
    """
    if not ANTHROPIC_API_KEY:
        return None

    i = summary_inputs or {}
    data_lines = []
    if i.get("campaign_name"):        data_lines.append(f"Campaign: {i['campaign_name']}")
    if i.get("client"):               data_lines.append(f"Client: {i['client']}")
    if i.get("markets"):              data_lines.append(f"Markets: {i['markets']}")
    if i.get("live_dates"):           data_lines.append(f"Live dates: {i['live_dates']}")
    if i.get("delivered_impressions"):data_lines.append(f"Delivered impressions: {i['delivered_impressions']}")
    if i.get("ctr"):                  data_lines.append(f"CTR: {i['ctr']}")
    if i.get("engagement_rate"):      data_lines.append(f"Engagement rate: {i['engagement_rate']}")
    if i.get("vcr"):                  data_lines.append(f"VCR: {i['vcr']}")
    if i.get("on_screen_rate"):       data_lines.append(f"On-screen rate: {i['on_screen_rate']}")
    if i.get("budget"):               data_lines.append(f"Budget: {i['budget']}")
    if i.get("delivery_incl_av"):     data_lines.append(f"Delivery incl. AV: {i['delivery_incl_av']}")
    if i.get("added_value_worth"):    data_lines.append(f"Added value: {i['added_value_worth']}")

    for key, label in [
        ("top_ctr",             "Top CTR titles"),
        ("top_engagement_rate", "Top engagement titles"),
        ("top_vcr",             "Top VCR titles"),
    ]:
        items = i.get(key)
        if isinstance(items, list) and items:
            top = ", ".join(x.get("name", "") for x in items[:3] if x.get("name"))
            if top:
                data_lines.append(f"{label}: {top}")

    if not data_lines:
        return None

    prompt = (
        "You are writing the executive summary for a premium digital advertising "
        "post-campaign analysis (PCA) report. Write a single, fluent paragraph of "
        "3-4 sentences that a senior account manager would be proud to send to a client. "
        "The tone should be confident, human, and results-focused — not corporate or "
        "template-sounding. Highlight the standout metrics and top performers naturally. "
        "Keep the paragraph under 500 characters. "
        "Do not use bullet points, headers, or markdown. Output only the paragraph, nothing else.\n\n"
        "Campaign data:\n" + "\n".join(data_lines)
    )

    try:
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.post(
                "https://api.anthropic.com/v1/messages",
                headers={
                    "x-api-key": ANTHROPIC_API_KEY,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json",
                },
                json={
                    "model": "claude-haiku-4-5-20251001",
                    "max_tokens": 300,
                    "messages": [{"role": "user", "content": prompt}],
                },
            )
        resp.raise_for_status()
        data = resp.json()
        text = data["content"][0]["text"].strip()
        if not text:
            return None
        # Hard cap at 550 chars, cutting at a sentence boundary where possible
        return _truncate_to_sentence(text, EXEC_SUMMARY_MAX_CHARS)
    except Exception:
        traceback.print_exc()
        return None


# ----- Auth -----------------------------------------------------------------

GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET", "")
ALLOWED_EMAIL_DOMAIN = os.getenv("ALLOWED_EMAIL_DOMAIN", "")
SESSION_SECRET = os.getenv("SESSION_SECRET", "change-me-in-render-env-vars")
APP_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://pca-modular-builder-v12.onrender.com",
)

oauth = OAuth()
oauth.register(
    name="google",
    client_id=GOOGLE_CLIENT_ID,
    client_secret=GOOGLE_CLIENT_SECRET,
    server_metadata_url="https://accounts.google.com/.well-known/openid-configuration",
    client_kwargs={"scope": "openid email profile"},
)


def current_user(request: Request) -> Optional[Dict[str, Any]]:
    return request.session.get("user")


def require_user(request: Request) -> Dict[str, Any]:
    user = current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    return user


# ----- Router ---------------------------------------------------------------

router = APIRouter()
templates = Jinja2Templates(directory=str(WEB_TEMPLATES_DIR))


@router.get("/login")
async def login(request: Request):
    redirect_uri = f"{APP_BASE_URL}/auth/callback"
    return await oauth.google.authorize_redirect(request, redirect_uri)


@router.get("/auth/callback")
async def auth_callback(request: Request):
    try:
        token = await oauth.google.authorize_access_token(request)
    except OAuthError as err:
        return JSONResponse({"error": str(err)}, status_code=400)

    user_info = token.get("userinfo")
    if not user_info:
        return JSONResponse({"error": "No user info returned"}, status_code=400)

    email = (user_info.get("email") or "").lower()
    if ALLOWED_EMAIL_DOMAIN and not email.endswith("@" + ALLOWED_EMAIL_DOMAIN.lower()):
        return templates.TemplateResponse(
            request,
            "denied.html",
            {"email": email, "allowed_domain": ALLOWED_EMAIL_DOMAIN},
            status_code=403,
        )

    with db_connect() as conn:
        conn.execute(
            """INSERT OR IGNORE INTO users (email, name, picture, created_at)
               VALUES (?, ?, ?, ?)""",
            (email, user_info.get("name"), user_info.get("picture"), datetime.utcnow().isoformat()),
        )

    request.session["user"] = {
        "email": email,
        "name": user_info.get("name"),
        "picture": user_info.get("picture"),
    }
    return RedirectResponse(url="/app")


@router.get("/logout")
async def logout(request: Request):
    request.session.clear()
    return RedirectResponse(url="/web")


# ----- Pages ----------------------------------------------------------------


@router.get("/web")
async def web_landing(request: Request):
    if current_user(request):
        return RedirectResponse(url="/app")
    return templates.TemplateResponse(request, "landing.html")


@router.get("/app")
async def app_page(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse(url="/web")
    return templates.TemplateResponse(request, "app.html", {"user": user})


@router.get("/history")
async def history_page(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse(url="/web")

    with db_connect() as conn:
        rows = conn.execute(
            """SELECT run_id, created_at, eoc_filename, output_type,
                      one_pager_filename, deck_filename, status
               FROM runs WHERE user_email = ?
               ORDER BY created_at DESC LIMIT 100""",
            (user["email"],),
        ).fetchall()

    return templates.TemplateResponse(
        request,
        "history.html",
        {
            "user": user,
            "runs": [dict(r) for r in rows],
            "retention_days": FILE_RETENTION_DAYS,
        },
    )


# ----- API ------------------------------------------------------------------


@router.get("/api/sections")
async def api_list_sections(user=Depends(require_user)):
    from main import load_section_registry, APP_VERSION

    registry = load_section_registry()
    sections = []
    for section_id, meta in registry.items():
        sections.append({
            "section_id": section_id,
            "label": meta.get("label", section_id),
            "required": meta.get("required", False),
            "default_order": meta.get("default_order", 999),
        })
    sections.sort(key=lambda x: x["default_order"])
    return {"app_version": APP_VERSION, "sections": sections}


@router.post("/api/validate")
async def api_validate(request: Request, eoc_file: UploadFile = File(...)):
    require_user(request)

    from main import (
        read_uploaded_eoc,
        build_mapped_values,
        compact_parsed_result,
        APP_VERSION,
    )

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        safe_name = (eoc_file.filename or "uploaded.xlsx").replace("/", "_").replace("\\", "_")
        eoc_path = tmp_dir / safe_name
        eoc_path.write_bytes(await eoc_file.read())

        try:
            sheets = read_uploaded_eoc(eoc_path)
            parsed = build_mapped_values(sheets, filename=eoc_path.name)
        except HTTPException:
            raise
        except Exception as exc:
            traceback.print_exc()
            raise HTTPException(status_code=400, detail=f"Validation failed: {exc}")

    compact = compact_parsed_result(parsed)
    summary_inputs = compact.get("summary_inputs") or {}

    # Claude API exec summary — async, capped at 550 chars, None on failure
    claude_summary = await _claude_exec_summary(summary_inputs)

    return JSONResponse({
        "status": "validated",
        "app_version": APP_VERSION,
        "filename": eoc_file.filename,
        "claude_exec_summary": claude_summary,
        **compact,
    })


@router.post("/api/generate")
async def api_generate(
    request: Request,
    eoc_file: UploadFile = File(...),
    output_type: str = Form(...),
    selected_sections: str = Form("[]"),
    deck_mode: str = Form("matching"),
    custom_deck_sections: str = Form("[]"),
    exec_summary: str = Form(""),
):
    user = require_user(request)

    from main import (
        read_uploaded_eoc,
        build_mapped_values,
        compact_parsed_result,
        resolve_exec_summary_for_ppt,
        build_grouped_stacked_modular_ppt,
        build_filtered_slide_deck_ppt,
        SlideDeckFromEocRequest,
    )

    try:
        section_list = json.loads(selected_sections) if selected_sections else []
        custom_list = json.loads(custom_deck_sections) if custom_deck_sections else []
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Invalid JSON in section fields")

    run_id = str(uuid.uuid4())
    one_pager_filename: Optional[str] = None
    deck_filename: Optional[str] = None
    summary_blob: Dict[str, Any] = {}
    campaign_name = ""

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        safe_name = (eoc_file.filename or "uploaded.xlsx").replace("/", "_").replace("\\", "_")
        eoc_path = tmp_dir / safe_name
        eoc_path.write_bytes(await eoc_file.read())

        try:
            sheets = read_uploaded_eoc(eoc_path)
            parsed = build_mapped_values(sheets, filename=eoc_path.name)
        except Exception as exc:
            traceback.print_exc()
            _record_failed_run(run_id, user["email"], eoc_file.filename, output_type)
            raise HTTPException(status_code=400, detail=f"Could not read EOC: {exc}")

        mapped_values = parsed["mapped_values"]
        mapped_values["EXEC_SUMMARY"] = resolve_exec_summary_for_ppt(exec_summary, mapped_values)
        summary_blob = compact_parsed_result(parsed)

        campaign_name = (
            (summary_blob.get("summary_inputs") or {}).get("campaign_name")
            or mapped_values.get("CAMPAIGN_NAME")
            or "PCA"
        )

        try:
            if output_type in ("one_pager", "both"):
                op_result = build_grouped_stacked_modular_ppt(
                    selected_sections=section_list,
                    placeholder_values=mapped_values,
                )
                original = Path(op_result["filename"]).name
                new_name = _smart_filename(campaign_name, "One_Pager")
                one_pager_filename = _rename_generated_file(original, new_name)

            if output_type in ("slide_deck", "both"):
                deck_request = SlideDeckFromEocRequest(
                    openaiFileIdRefs=[],
                    selected_sections=section_list,
                    one_pager_sections=section_list,
                    custom_deck_sections=custom_list,
                    deck_mode=deck_mode,
                    exec_summary=exec_summary,
                )
                deck_result = build_filtered_slide_deck_ppt(
                    request=deck_request,
                    placeholder_values=mapped_values,
                )
                original = Path(deck_result["filename"]).name
                new_name = _smart_filename(campaign_name, "Slide_Deck")
                deck_filename = _rename_generated_file(original, new_name)

        except Exception as exc:
            traceback.print_exc()
            _record_failed_run(run_id, user["email"], eoc_file.filename, output_type)
            raise HTTPException(status_code=500, detail=f"Generation failed: {exc}")

    with db_connect() as conn:
        conn.execute(
            """INSERT INTO runs
               (run_id, user_email, created_at, eoc_filename, output_type,
                selected_sections, deck_mode, exec_summary,
                one_pager_filename, deck_filename, summary_json, status)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                run_id,
                user["email"],
                datetime.utcnow().isoformat(),
                eoc_file.filename,
                output_type,
                json.dumps(section_list),
                deck_mode,
                exec_summary,
                one_pager_filename,
                deck_filename,
                json.dumps(summary_blob),
                "completed",
            ),
        )

    return {
        "success": True,
        "run_id": run_id,
        "one_pager_url": f"/files/{one_pager_filename}" if one_pager_filename else None,
        "deck_url": f"/files/{deck_filename}" if deck_filename else None,
        "summary": summary_blob,
    }


def _record_failed_run(run_id: str, email: str, filename: Optional[str], output_type: str) -> None:
    try:
        with db_connect() as conn:
            conn.execute(
                """INSERT INTO runs
                   (run_id, user_email, created_at, eoc_filename, output_type,
                    selected_sections, deck_mode, exec_summary,
                    one_pager_filename, deck_filename, summary_json, status)
                   VALUES (?, ?, ?, '[]', '', '', NULL, NULL, '{}', 'failed')""",
                (run_id, email, datetime.utcnow().isoformat(), filename, output_type),
            )
    except Exception:
        pass


@router.get("/api/runs/{run_id}")
async def api_get_run(run_id: str, user=Depends(require_user)):
    with db_connect() as conn:
        row = conn.execute(
            "SELECT * FROM runs WHERE run_id = ? AND user_email = ?",
            (run_id, user["email"]),
        ).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Run not found")
    return dict(row)


# ----- Registration ---------------------------------------------------------


def register_web_routes(app: FastAPI) -> None:
    app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET)
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
    app.include_router(router)

    @app.middleware("http")
    async def _cleanup_starter_middleware(request, call_next):
        _ensure_cleanup_task_running()
        return await call_next(request)
