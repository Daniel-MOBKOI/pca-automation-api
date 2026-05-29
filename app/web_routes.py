"""
web_routes.py
==============
Web-app endpoints for the PCA Automation Generator.

Lives at: app/web_routes.py (alongside app/main.py)

v7 changes:
- Split admin functionality into its own dedicated /admin route.
  /history is back to being a simple per-user view (matches v5 behaviour).
  /admin shows the team-wide per-user summary with expandable rows.
- Non-admins who try to hit /admin directly get silently redirected to
  /history (defence in depth — the UI also hides the link).
- New admin.html template.

Earlier versions:
- v6: admin-aware /history (replaced by this version's clean split)
- v5: /logout redirects to /web, smart filename renaming
- v4: SessionMiddleware cleanup task started by middleware
- v3: persistent SQLite + per-user history
- v2: Google OAuth Internal app
- v1: original web wizard endpoints
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
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

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
    if not s:
        return "PCA"
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    cleaned = re.sub(r"_+", "_", cleaned)
    return cleaned[:max_len] or "PCA"


def _smart_filename(campaign_name: str, doc_type: str) -> str:
    date_part = datetime.utcnow().strftime("%d%m%y")
    name = _safe_filename_part(campaign_name or "PCA")
    return f"{name}_{doc_type}_{date_part}.pptx"


def _rename_generated_file(original_filename: str, new_filename: str) -> str:
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


# ----- Auth + admin support -------------------------------------------------

from authlib.integrations.starlette_client import OAuth, OAuthError
from starlette.middleware.sessions import SessionMiddleware

GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET", "")
ALLOWED_EMAIL_DOMAIN = os.getenv("ALLOWED_EMAIL_DOMAIN", "")
SESSION_SECRET = os.getenv("SESSION_SECRET", "change-me-in-render-env-vars")
APP_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://pca-modular-builder-v12.onrender.com",
)

# ADMIN_EMAILS is a comma-separated list of email addresses, e.g.
#   ADMIN_EMAILS=daniel.crittenden@mobkoi.com,jane.doe@mobkoi.com
# Comparison is case-insensitive; whitespace and empty entries are ignored.
ADMIN_EMAILS = {
    e.strip().lower()
    for e in os.getenv("ADMIN_EMAILS", "").split(",")
    if e.strip()
}


def is_admin(user: Optional[Dict[str, Any]]) -> bool:
    if not user:
        return False
    email = (user.get("email") or "").lower()
    return email in ADMIN_EMAILS


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


# ----- Template context injection -------------------------------------------
#
# Every page that renders base.html needs to know whether the current user
# is an admin (so the nav can conditionally show the "Admin" link). Rather
# than passing is_admin manually on every TemplateResponse, we provide a
# small helper that all routes use.

def _page_context(request: Request, **extra) -> Dict[str, Any]:
    user = current_user(request)
    ctx: Dict[str, Any] = {
        "user": user,
        "is_admin": is_admin(user),
    }
    ctx.update(extra)
    return ctx


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
    return templates.TemplateResponse(request, "landing.html", _page_context(request))


@router.get("/app")
async def app_page(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse(url="/web")
    return templates.TemplateResponse(request, "app.html", _page_context(request))


@router.get("/history")
async def history_page(request: Request):
    """
    Always shows the logged-in user's OWN PCAs only.
    Admins use /admin for the team-wide view.
    """
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
        _page_context(
            request,
            runs=[dict(r) for r in rows],
            retention_days=FILE_RETENTION_DAYS,
        ),
    )


# ----- Admin page -----------------------------------------------------------


def _relative_time(iso_str: str) -> str:
    """
    Convert an ISO timestamp to a friendly relative string:
    'just now', '5 min ago', '2 hours ago', 'yesterday', '3 days ago',
    or fallback to the date for anything older than 14 days.
    """
    if not iso_str:
        return "—"
    try:
        dt = datetime.fromisoformat(iso_str.replace("Z", ""))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        now = datetime.now(timezone.utc)
        delta = now - dt
        seconds = int(delta.total_seconds())
        if seconds < 60:
            return "just now"
        minutes = seconds // 60
        if minutes < 60:
            return f"{minutes} min ago"
        hours = minutes // 60
        if hours < 24:
            return f"{hours} hour{'s' if hours != 1 else ''} ago"
        days = hours // 24
        if days == 1:
            return "yesterday"
        if days < 14:
            return f"{days} days ago"
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return iso_str[:10] if len(iso_str) >= 10 else iso_str


# Friendly labels for output_type values stored in the DB.
_OUTPUT_TYPE_LABELS = {
    "one_pager": "One Pager",
    "slide_deck": "Slide Deck",
    "both": "One Pager + Slide Deck",
}


def _build_user_summaries(all_runs: List[sqlite3.Row]) -> List[Dict[str, Any]]:
    """
    Group a flat list of run rows by user and compute per-user summary stats.

    Each summary dict contains:
        email, name, picture, total_pcas, last_active (relative),
        last_active_iso (raw, used for sorting), most_common_type,
        first_used (date only), and runs (full list for the expanded row).

    Sorted by total_pcas DESC, then most recent activity as tiebreaker,
    so the most prolific users surface at the top.
    """
    grouped: Dict[str, List[sqlite3.Row]] = defaultdict(list)
    for row in all_runs:
        grouped[row["user_email"]].append(row)

    user_meta: Dict[str, Dict[str, Any]] = {}
    if grouped:
        with db_connect() as conn:
            placeholders = ",".join("?" * len(grouped))
            rows = conn.execute(
                f"SELECT email, name, picture, created_at FROM users WHERE email IN ({placeholders})",
                tuple(grouped.keys()),
            ).fetchall()
            for r in rows:
                user_meta[r["email"]] = dict(r)

    summaries = []
    for email, runs in grouped.items():
        type_counter = Counter(r["output_type"] for r in runs if r["output_type"])
        most_common_raw = type_counter.most_common(1)[0][0] if type_counter else None
        most_common = _OUTPUT_TYPE_LABELS.get(most_common_raw, most_common_raw or "—")

        last_active_iso = runs[0]["created_at"] if runs else None
        first_used_iso = runs[-1]["created_at"] if runs else None

        meta = user_meta.get(email, {})
        summaries.append({
            "email": email,
            "name": meta.get("name") or email,
            "picture": meta.get("picture"),
            "total_pcas": len(runs),
            "last_active": _relative_time(last_active_iso),
            "last_active_iso": last_active_iso,
            "most_common_type": most_common,
            "first_used": first_used_iso[:10] if first_used_iso else "—",
            "runs": [dict(r) for r in runs],
        })

    summaries.sort(
        key=lambda s: (s["total_pcas"], s["last_active_iso"] or ""),
        reverse=True,
    )
    return summaries


@router.get("/admin")
async def admin_page(request: Request):
    """
    Admin-only page showing per-user PCA activity across the whole team.
    Non-admins (and unauthenticated visitors) get silently redirected to
    /history — no scary error page, just sent to a place they can use.
    """
    user = current_user(request)
    if not user:
        return RedirectResponse(url="/web")
    if not is_admin(user):
        return RedirectResponse(url="/history")

    with db_connect() as conn:
        rows = conn.execute(
            """SELECT run_id, user_email, created_at, eoc_filename, output_type,
                      one_pager_filename, deck_filename, status
               FROM runs
               ORDER BY created_at DESC
               LIMIT 1000"""
        ).fetchall()

    user_summaries = _build_user_summaries(rows)

    total_pcas = len(rows)
    active_user_count = len(user_summaries)

    week_ago = (datetime.utcnow() - timedelta(days=7)).isoformat()
    recent_emails = {r["user_email"] for r in rows if r["created_at"] >= week_ago}
    active_this_week = len(recent_emails)
    pcas_this_week = sum(1 for r in rows if r["created_at"] >= week_ago)

    return templates.TemplateResponse(
        request,
        "admin.html",
        _page_context(
            request,
            user_summaries=user_summaries,
            retention_days=FILE_RETENTION_DAYS,
            headline_stats={
                "total_pcas": total_pcas,
                "active_users_total": active_user_count,
                "active_users_week": active_this_week,
                "pcas_this_week": pcas_this_week,
            },
        ),
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
    app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET)
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
    app.include_router(router)

    @app.middleware("http")
    async def _cleanup_starter_middleware(request, call_next):
        _ensure_cleanup_task_running()
        return await call_next(request)
