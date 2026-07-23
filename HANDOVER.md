# PCA Generator — Project Handover

> **How this file works:** This is the living source of truth for the project, kept in the repo itself (not just in Claude project knowledge). Claude reads it at the start of every session and updates it after any change we make together, so we never work from a stale snapshot again. This includes working rules and preferences Daniel gives along the way (not just code/feature state) — Claude folds these into "How Daniel likes to work" / "Lessons learned" as they come up, rather than relying on the Claude.ai project's memory. As of 23 July 2026, the old project knowledge files (code snapshots, `HANDOVER_v3.md`) are being deleted from the Claude.ai project since this file supersedes them.

---

## TL;DR

Daniel Crittenden, Mobkoi (digital media agency). The **PCA Generator** turns end-of-campaign Excel files (EOCs) into branded PowerPoint deliverables (One Pager and/or Slide Deck). **Live, in production, rolled out to the team.** Phase 1 launch went well — app is being used and well received. We're now in snag/tweak mode based on real usage.

**Live URL:** https://pca-modular-builder-v12.onrender.com
**Repo:** github.com/Daniel-MOBKOI/pca-automation-api
**Local working copy:** Daniel now works locally (this folder), pushes to GitHub, Render auto-deploys from `main` on push.
**Cost:** ~$7.25/mo Render + ~$9/mo Anthropic API (moderate use) ≈ **~$16/mo all-in**

---

## Current repo state (verified 23 July 2026)

- **Latest commit:** `331f2d6` ("Update .DS_Store"), on `Daniel-MOBKOI/pca-automation-api` main, working tree clean.
- **GitHub Desktop / fork gotcha (resolved 23 July 2026):** when Daniel connected this local folder to GitHub Desktop, Desktop's signed-in account (`danjcDesign`) didn't have write access to `Daniel-MOBKOI/pca-automation-api` (the actual Render deploy source), so it silently auto-forked the repo to `danjcDesign/pca-automation-api` and pushed there instead. Two commits (this file's addition + a `.DS_Store` update) went to the fork, not the org repo, and briefly went "missing" from Daniel's perspective viewing the org repo on GitHub. Fixed by: adding `danjcDesign` as a collaborator on `Daniel-MOBKOI/pca-automation-api`, then repointing the local `origin` remote from the fork to the org repo directly (fork kept around under remote name `fork` for reference, not actively used). **Lesson: if GitHub Desktop ever shows "publish fork" language or a repo appears out of sync with what's on GitHub, check `git remote -v` first — Desktop silently forks when the signed-in account lacks write access to the original remote.**
- **No commits since 4 June** — a ~7 week gap while the team used the app in production, which lines up with "phase one launch went well."
- This file (`HANDOVER.md`) did not previously exist in the repo — the handover doc only lived in Claude project knowledge (`HANDOVER_v3.md`) and had drifted out of date relative to what's actually in `main`. That version *undersold* what's shipped: it still listed Japanese localisation and the Round 1 performance fix as in-progress/next-step, but both are already merged into `main` and live.

### What's actually in `main` right now

- ✅ **Round 1 performance fix** — `app/web_routes.py` is v10. `/api/generate` is a sync `def` (runs in FastAPI's worker thread pool); `/api/validate` stays `async` but offloads the heavy EOC parse to a worker thread via `run_in_threadpool` (`_parse_eoc_for_validation` helper). Fixes one user's upload freezing page loads for everyone else on the single-worker instance.
- ✅ **Japanese localisation** — fully coded and merged:
  - `app/main.py`: `template_path_for()` helper picks `*_ja.pptx` master when language is `ja`, falls back to English master otherwise. Both `build_one_pager` / slide-deck generation functions accept an optional `template_path` override.
  - `app/templates/modular_sections_master_ja.pptx` and `slide_deck_master_ja.pptx` exist in the repo (Noto Sans JP masters).
  - `app/web_templates/app.html`: language selector UI (flag icons, EN/JA chips) next to the output-type picker; per-EOC `language` state; `setLanguage()` switches language, triggers exec-summary regeneration, and reverts on cancel via `state.langRevert`; typewriter animation changed from character-by-character to word-by-word (needed since Japanese has no spaces, and reads more naturally in English too).
  - `/api/exec-summary` and `/api/generate` both now pass `language` through.

**Not yet confirmed:** whether this has been regression-tested against the live Render deployment, or whether Render has actually redeployed this commit. Verify with a screenshot before assuming any specific feature works, per the lesson below.

---

## Architecture

```
┌──────────────────────────────────────────────────────┐
│  pca-automation-api (FastAPI on Render)              │
│                                                       │
│  ┌──────────────────┐    ┌─────────────────────────┐ │
│  │ app/main.py      │    │ app/web_routes.py       │ │
│  │ (existing)       │    │ (added by us)           │ │
│  │                  │    │                         │ │
│  │ - EOC parsing    │◄───┤ - Web endpoints         │ │
│  │ - PPT generation │    │ - Google OAuth          │ │
│  │ - GPT-facing API │    │ - SQLite database       │ │
│  │ - JA template     │    │ - Auto cleanup          │ │
│  │   selection       │    │ - Smart file renaming   │ │
│  │                  │    │ - Claude API integration│ │
│  │                  │    │ - Admin role            │ │
│  └──────────────────┘    └─────────────────────────┘ │
│         ▲                            ▲                │
│         │                            │                │
│  ┌──────┴──────┐            ┌────────┴────────┐      │
│  │ Custom GPT  │            │ Web wizard      │      │
│  │ (existing)  │            │ (5-step flow)   │      │
│  └─────────────┘            └─────────────────┘      │
└──────────────────────────────────────────────────────┘
                              │
                ┌─────────────┼─────────────┐
                │             │             │
        ┌───────▼──────┐ ┌────▼─────┐ ┌─────▼──────┐
        │ Render disk  │ │ Google   │ │ Anthropic  │
        │ /var/data    │ │ OAuth    │ │ Claude API │
        │  ├─pca.db    │ │ (Internal│ │ (haiku-4-5)│
        │  └─generated/│ │  app)    │ │            │
        └──────────────┘ └──────────┘ └────────────┘
```

**Key principle:** the web app is purely additive to `main.py`. All existing GPT endpoints are unchanged.

---

## Repo structure

Render config: `Root Directory: app`, `Build: pip install -r ../requirements.txt`, `Start: uvicorn main:app --host 0.0.0.0 --port $PORT`

```
pca-automation-api/                  ← repo root (now also the local working copy)
├── HANDOVER.md                      ← this file — living source of truth
├── README.md
├── requirements.txt
├── app/                              ← Render's root dir
│   ├── main.py                       ← includes template_path_for() for JA masters
│   ├── web_routes.py                 ← v10 (concurrency fix + Claude API + admin + smart filenames)
│   ├── web_templates/
│   │   ├── base.html                 ← v6 (conditional Admin link in nav)
│   │   ├── landing.html              ← v3
│   │   ├── app.html                  ← language selector + word-by-word typewriter (version comment not present in file — treat version numbers here as approximate until confirmed)
│   │   ├── history.html              ← v8 (per-user, search & filter)
│   │   ├── admin.html                ← v2 (team activity dashboard)
│   │   └── denied.html               ← v2
│   ├── static/
│   │   ├── mobkoi-logo.webp
│   │   └── favicon.ico
│   ├── section_registry/             ← existing JSON registry
│   └── templates/                    ← PPTX masters, including *_ja.pptx variants
├── templates/                        ← more existing PPTX masters
└── assets/
```

---

## User journey

Five-step wizard at `/app` after Google sign-in:

1. **Upload** — drag-and-drop one to three EOC `.xlsx` files (batch upload supported)
2. **Review** — bento tiles (Campaign Overview / Performance Metrics / Delivery & Spend / Top Performing Titles / Top Performing Markets). Expandable "See Full Validation Report".
3. **Customise** — per EOC (tabs): output type (One Pager / Full Deck / Both), language (EN/JA), section picks, edit/regenerate exec summary (Claude-written, "✨ Claude" pill).
4. **Generate** — animated progress
5. **Done** — download cards with smart filenames like `Prada_Eyewear_FW25_One_Pager_290526.pptx`, plus "Built in X:XX" timing

Plus `/history` (per-user, search + date filter) and `/admin` (team activity dashboard, admins only).

---

## Tech stack

- **Backend:** FastAPI + Python 3.11
- **Auth:** Google Workspace SSO via authlib, restricted to `@mobkoi.com` Internal app
- **Storage:** SQLite on Render persistent disk + ephemeral generated PPTX files (30-day auto-cleanup)
- **AI:** Anthropic Claude API (`claude-haiku-4-5`) via httpx for exec summary generation
- **Frontend:** Server-rendered Jinja2 + Tailwind via CDN + vanilla JS (no build step)
- **Deploy:** Render auto-deploys from `main` branch on push

---

## Environment variables on Render

```
GOOGLE_CLIENT_ID         <set>
GOOGLE_CLIENT_SECRET     <set>
ALLOWED_EMAIL_DOMAIN     mobkoi.com
SESSION_SECRET           <auto-generated random>
PUBLIC_BASE_URL          https://pca-modular-builder-v12.onrender.com
PCA_DB_PATH              /var/data/pca.db
GENERATED_FILES_DIR      /var/data/generated
FILE_RETENTION_DAYS      30 (optional; defaults to 30)
ANTHROPIC_API_KEY        <set — required for Claude exec summaries>
ADMIN_EMAILS             daniel.crittenden@mobkoi.com (comma-separated for multiple)
```

Render disk attached: 1GB at `/var/data` (~£0.25/mo). Instance tier: Standard (upgraded from Starter as part of the Round 1 performance fix).

---

## Claude API integration

**Model:** `claude-haiku-4-5`. **Cost:** ~$0.001/call, ~$9/mo at 10 users × 3 PCAs/day.

**Flow:** `/api/validate` parses EOC → calls Claude with structured campaign data → returns a 3-4 sentence paragraph in Mobkoi's voice → frontend shows it with "✨ Claude" pill → Regenerate calls `/api/exec-summary` → silent fallback to JS templates on any Claude failure.

**System prompt** (in `_claude_exec_summary()`): "You are writing the executive summary for a premium digital advertising post-campaign analysis (PCA) report. Write a single, fluent paragraph of 3-4 sentences that a senior account manager would be proud to send to a client. The tone should be confident, human, and results-focused — not corporate or template-sounding. Highlight the standout metrics and top performers naturally. Keep the paragraph under 500 characters. Do not use bullet points, headers, or markdown. Output only the paragraph, nothing else."

**Debugging:** `[claude] ANTHROPIC_API_KEY is not set; skipping Claude summary` logs to Render if the key goes missing — early warning for silent regressions.

---

## Design system

- **Typography (Review screen):** Section title 13px semibold Title Case grey; tile label 10px ALL CAPS lightest grey; tile value 15px semibold black.
- **Capitalisation rule:** two-word headings/titles get each word capitalised — "One Pager", "Slide Deck", "EOC File", "Build Another".
- **Live Dates display:** `14 Oct – 02 Nov 2025` (abbreviated months, en-dash) in Review only; underlying value stays full format for PPT injection.
- **Colour palette:** accent blue `#2563eb`, surfaces `#ffffff`/`#f8fafc`, text `#0f172a`/`#475569`/`#94a3b8`. Inter font; Noto Sans JP for Japanese output.

---

## What's shipped (chronological)

**Initial build:** Render disk + env vars; Google OAuth (Internal, `@mobkoi.com`); five-step wizard; per-user history; 30-day auto cleanup; bento-tile Review screen; Top Titles/Markets; conditional Customise screen; "One Pager" capitalisation.

**Round 1 (bugs + quick wins):** Header Sign In → `/login`; Sign Out → `/web`; History Downloads column alignment; smart download filenames; inline edit-warning modal before Regenerate overwrites edits; "Not enough data" badges; silent build timer; History page polish.

**Round 2:** History search by EOC filename + date range filter.

**Round 3 (architectural):** Batch upload (up to 3 EOCs, tabbed wizard); Claude API exec summary + `/api/exec-summary` endpoint + "✨ Claude" pill; silent fallback to JS templates on Claude failure.

**Round 4 (admin dashboard):** `ADMIN_EMAILS` env var + `is_admin()`; dedicated `/admin` page; per-user summary table with expandable rows; headline stats; non-admins silently redirected from `/admin` to `/history`.

**Round 5 (performance + localisation, committed 3–4 June, not yet listed in Claude's project knowledge until this update):**
- Standard tier upgrade; `/api/generate` and `/api/validate` heavy work moved off the event loop.
- Japanese localisation: JA master templates, language selector UI, per-EOC language toggle with exec-summary regen and revert-on-cancel, word-by-word typewriter animation.

**Round 6 (Phase 1 feedback batch, 23 July 2026 — committed locally, not yet pushed/deployed):**
- ✅ **AV delivery decimal fix**: `format_percent()` in `main.py` now takes an `always_scale` flag; `DELIVERY_WITH_AV_PERCENT` uses it so over-100% values (e.g. raw 1.54 for 154% over-delivery) render as "154%" instead of "1.54%". Root cause was the old `if num <= 1: *100` heuristic misreading an already->1 fraction as "already scaled."
- ✅ **On-Screen → MRC Viewability fallback rename**: when the EOC has no true on-screen column and falls back to a dedicated MRC Viewability column, the label now says "MRC Viewability" instead of misrepresenting it as "On-Screen Rate" — in the Review tile, Claude exec-summary input, validation text, and the English Slide Deck PPT (label was static text next to the value token; turned into a matching `{{PERFORMANCE_ON_SCREEN_LABEL}}` token). **Not yet done:** the Japanese Slide Deck template's label ("オンスクリーン") is still hand-translated static text — needs Daniel to supply the correct Japanese phrase for "MRC Viewability" before that template gets the same treatment. One Pager template doesn't include this metric at all, so nothing needed there.
- ✅ **Fixed a real bug found while scoping the currency feature: Japanese localisation was not actually wired end-to-end.** `app.html` sent `language` to both `/api/exec-summary` (JSON body) and `/api/generate` (form field), but `web_routes.py` silently ignored it in both places — `_claude_exec_summary()` never got a language hint, and `/api/generate` never computed/passed a `template_path` to the PPT builders (even though `main.py`'s `template_path_for()` helper and both builder functions were ready to accept one). Practical effect: selecting Japanese in the UI likely produced an English-template deck with an English-toned summary. Fixed: `_claude_exec_summary()` now takes a `language` param and asks Claude for natural Japanese when `ja`; `_truncate_to_sentence()` now also recognises Japanese sentence punctuation (。！？) as a truncation boundary; `/api/exec-summary` reads `language` from the request body; `/api/generate` now declares a `language` form field and passes `template_path_for(MODULAR_TEMPLATE_PATH, language)` / `template_path_for(SLIDE_DECK_TEMPLATE_PATH, language)` into the one-pager/slide-deck builders. **This needs a real end-to-end test against a Japanese EOC before trusting it in front of the team** — verified the template-selection logic and placeholder substitution mechanics in isolation, but haven't run the live app.
- ✅ **Currency detection + Customise-tab Currency dropdown**: added `detect_currency_symbol()` in `main.py` — it reads the raw (pre-`safe_number`-stripped) spend/added-value-worth cell text for a known symbol (£/$/€/¥). `build_mapped_values()` now takes an optional `currency_override` param; resolution order is explicit override → detected symbol → `$` default. Both `CAMPAIGN_BUDGET` and `ADDED_VALUE_WORTH` now format with the resolved symbol instead of a hardcoded `€`. `/api/validate`'s `summary_inputs` now includes `campaign_currency` so the Review-screen tiles and the Customise-tab dropdown both reflect the detected currency automatically. Added a Currency chip selector (€/$/£/¥) to the Customise tab next to output type and language — pre-selected to the detected currency, user can override. `/api/generate` now accepts a `currency` form field and re-derives `mapped_values` with that override, so the generated One Pager/Slide Deck uses whichever currency the user picked. **Scope note:** the currency override only affects the generated PPT output (matches what Daniel asked for — "for the one pager or slides to use"); it does not retroactively reformat the Review-screen tiles or trigger an exec-summary regenerate the way the language toggle does, since Claude's summary prompt doesn't typically quote exact budget figures. Flag if you want it to also sync back to the Review tiles.

---

## Outstanding snag list

- [ ] Confirm the JA wiring fix above actually produces a Japanese-templated deck + Japanese-toned summary when tested live (was fixed based on code inspection, not yet verified against a real generate call).
- [ ] Get Daniel's preferred Japanese phrase for "MRC Viewability" so the JA Slide Deck template's fallback label can be updated to match the English one.
- [ ] All of Round 6 (decimal fix, MRC rename, JA wiring fix, currency) needs a real local/staging test before pushing to `main` — none of it has been run against the live app yet, only verified via isolated logic tests and syntax checks (couldn't run the full FastAPI app in this sandbox — its SQLite init fails outside Daniel's actual environment).

*(Empty items below awaiting further Phase 1 feedback from the team.)*

---

## Lessons learned

1. **Claude regression (Round 3):** the entire Claude integration was once silently removed from `web_routes.py` by an unrelated overwrite. Always verify a "working" feature with a screenshot before building on top of it. The `[claude] ANTHROPIC_API_KEY is not set` log line exists so this self-diagnoses in future.
2. **Don't just defer on model choice when stakes are real** — show a concrete side-by-side before picking Haiku vs Sonnet, don't just recommend.
3. **Admin access is data-driven** (`ADMIN_EMAILS` env var), not DB-driven — intentional, keep it simple until there's a real need for more roles.
4. **Project snapshots/knowledge caches drift** — this file existing in the repo (rather than only in Claude project knowledge) is the fix. Treat the repo as ground truth; update this file after every session.

---

## Phase 2 parking lot (not built)

**Worth doing eventually:** custom domain (`pca.mobkoi.com`); sharable expiring links; "Re-run with new EOC" history button; Slack notifications; in-app PPT thumbnail preview; brand/client memory; EOC validation hints (e.g. CTR > 5%, missing budget).

**Bigger ideas:** comparison mode (diff two past PCAs); AI-suggested section selection; save-as-template; mobile view check.

**Don't build proactively:** user roles beyond admin; favourites/tags on history; multi-tenant SaaS; charts on admin dashboard; section-popularity analytics.

---

## How Daniel likes to work

- Now works **locally** in this folder (previously GitHub web UI only) and pushes to GitHub himself; Render auto-deploys from `main`.
- Wants honest feedback, not cheerleading — call out overengineering.
- Iterates on annotated screenshots with specific design notes.
- Pragmatic — ships working over perfect, happy to defer features.
- Wants understanding played back before any code is written; use clarifying questions rather than guessing on ambiguous scope.
- Prefers small, reviewable, reversible changes over big-bang deploys.
- Treats every session as fresh, but now this file removes the need to re-paste context — **read this file at the start of every session.**

---

## How to keep working

1. **Read this file first**, every session — it is the ground truth, not the Claude project knowledge cache.
2. Listen for new feedback/bugs (Daniel often pastes screenshots with annotations).
3. Play back understanding before coding — list the specific changes, ask clarifying questions.
4. Small, reversible commits; verify each step before moving to the next.
5. Keep the **Outstanding snag list** section current — add items as Daniel reports them, check them off as they ship.
6. **Update this file after every session/change** — new shipped items, new snags, new commit hashes, any drift discovered.
7. Since Daniel now works locally with git, commits can happen directly in this folder — confirm with Daniel before pushing to `origin/main` (Render auto-deploys on push, so pushes go live).

---

## Latest verified state

- **Commit:** `41c8382` (4 June 2026)
- **`app/web_routes.py`:** v10
- **`app/web_templates/base.html`:** v6
- **`app/web_templates/history.html`:** v8
- **`app/web_templates/admin.html`:** v2
- **`app/web_templates/landing.html`:** v3
- **`app/web_templates/denied.html`:** v2
- **`app/web_templates/app.html`:** language selector + word-by-word typewriter shipped; no version comment in file to confirm exact version number
- **`app/main.py`:** `template_path_for()` JA helper shipped; otherwise untouched aside from the 2-line web-routes hook
