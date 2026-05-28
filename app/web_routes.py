"""
web_routes.py
==============
Web-app endpoints for the PCA Automation Generator.

Lives at: app/web_routes.py (alongside app/main.py)

This module adds a parallel set of endpoints designed for browser-based
multipart file uploads (instead of OpenAI Actions' openaiFileIdRefs pattern).

It REUSES all existing logic from main.py - no main.py functions are
modified, only imported.

The existing GPT-facing endpoints in main.py remain untouched. Both
surfaces can run side by side.

To wire this in, add to the bottom of main.py:
    from web_routes import register_web_routes
    register_web_routes(app)
"""

from __future__ import annotations

import json
import os
import sqlite3
import tempfile
import traceback
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

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


# ----- Paths ----------------------------------------------------------------
# This file lives in app/, same place as main.py
APP_DIR = Path(__file__).resolve().parent
WEB_TEMPLATES_DIR = APP_DIR / "web_templates"
STATIC_DIR = APP_DIR / "static"
STATIC_DIR.mkdir(parents=True, exist_ok=True)

# DB lives on the Render disk (persistent). Falls back to /tmp locally.
DB_PATH = Path(os.getenv("PCA_DB_PATH", APP_DIR / "pca.db"))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


# ----- Database (SQLite) ----------------------------------------------------


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


# ----- Auth (Google OAuth) --------------------------------------------------

from authlib.integrations.starlette_client import OAuth, OAuthError
from starlette.middleware.sessions import SessionMiddleware

GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET", "")
ALLOWED_EMAIL_DOMAIN = os.getenv("ALLOWED_EMAIL_DOMAIN", "")  # e.g. "mobkoi.com"
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
            "denied.html",
            {"request": request, "email": email, "allowed_domain": ALLOWED_EMAIL_DOMAIN},
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
    return RedirectResponse(url="/")


# ----- Pages ----------------------------------------------------------------


@router.get("/web")
async def web_landing(request: Request):
    """
    Web app landing page. Lives at /web instead of / because the existing
    main.py already owns / for its JSON health response (used by the GPT).
    """
    if current_user(request):
        return RedirectResponse(url="/app")
    return templates.TemplateResponse("landing.html", {"request": request})


@router.get("/app")
async def app_page(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse(url="/web")
    return templates.TemplateResponse("app.html", {"request": request, "user": user})


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
        "history.html",
        {"request": request, "user": user, "runs": [dict(r) for r in rows]},
    )


# ----- API used by the wizard frontend --------------------------------------
#
# All imports of main.py happen INSIDE function bodies to avoid any
# circular-import issues at module load time.


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
    """Accepts a direct EOC upload and runs validation."""
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

    return JSONResponse({
        "status": "validated",
        "app_version": APP_VERSION,
        "filename": eoc_file.filename,
        **compact_parsed_result(parsed),
    })


@router.post("/api/generate")
async def api_generate(
    request: Request,
    eoc_file: UploadFile = File(...),
    output_type: str = Form(...),                  # 'one_pager' | 'slide_deck' | 'both'
    selected_sections: str = Form("[]"),
    deck_mode: str = Form("matching"),
    custom_deck_sections: str = Form("[]"),
    exec_summary: str = Form(""),
):
    """Accepts EOC + choices, runs the existing pipeline, stores the result."""
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

        try:
            if output_type in ("one_pager", "both"):
                op_result = build_grouped_stacked_modular_ppt(
                    selected_sections=section_list,
                    placeholder_values=mapped_values,
                )
                one_pager_filename = Path(op_result["filename"]).name

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
                deck_filename = Path(deck_result["filename"]).name

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
                   VALUES (?, ?, ?, ?, ?, '[]', '', '', NULL, NULL, '{}', 'failed')""",
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
    """Wire this module into the existing FastAPI app."""
    app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET)
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
    app.include_router(router)
