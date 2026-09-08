# PCA Generator — Project Handover

> **How this file works:** This is the living source of truth for the project, kept in the repo itself (not just in Claude project knowledge). Claude reads it at the start of every session and updates it after any change we make together, so we never work from a stale snapshot again. This includes working rules and preferences Daniel gives along the way (not just code/feature state) — Claude folds these into "How Daniel likes to work" / "Lessons learned" as they come up, rather than relying on the Claude.ai project's memory. As of 23 July 2026, the old project knowledge files (code snapshots, `HANDOVER_v3.md`) are being deleted from the Claude.ai project since this file supersedes them.

---

## TL;DR

Daniel Crittenden, Mobkoi (digital media agency). The **PCA Generator** turns end-of-campaign Excel files (EOCs) into branded PowerPoint deliverables (One Pager and/or Slide Deck). **Live, in production, rolled out to the team.** Phase 1 launch went well — app is being used and well received. We're now in snag/tweak mode based on real usage.

**Live URL:** https://pca.mobkoi.com (custom domain, added 28 July 2026; old `pca-modular-builder-v12.onrender.com` link still works but 307-redirects to the canonical domain — see Round 7 below)
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
PUBLIC_BASE_URL          https://pca.mobkoi.com   (updated 28 July 2026 — was the onrender.com URL; this also feeds download_url + the OAuth redirect_uri, so it must always match whichever domain the team is meant to log in on)
PCA_DB_PATH              /var/data/pca.db
GENERATED_FILES_DIR      /var/data/generated
FILE_RETENTION_DAYS      14 (explicitly set 12 Aug 2026 — previously unset in Render, so was silently running on the code default of 30; see Round 8)
ANTHROPIC_API_KEY        <set — required for Claude exec summaries>
ADMIN_EMAILS             daniel.crittenden@mobkoi.com (comma-separated for multiple)
```

Render disk attached: 5GB at `/var/data` (resized from 1GB on 12 Aug 2026 — see Round 8). Instance tier: Standard (upgraded from Starter as part of the Round 1 performance fix).

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
- ✅ **Currency detection + Customise-tab Currency dropdown**: added `detect_currency_symbol()` in `main.py` — it reads the raw (pre-`safe_number`-stripped) spend/added-value-worth cell text for a known symbol (£/$/€/¥). `build_mapped_values()` now takes an optional `currency_override` param; resolution order is explicit override → detected symbol → `$` default. Both `CAMPAIGN_BUDGET` and `ADDED_VALUE_WORTH` now format with the resolved symbol instead of a hardcoded `€`. `/api/validate`'s `summary_inputs` now includes `campaign_currency` so the Review-screen tiles and the Customise-tab dropdown both reflect the detected currency automatically. Added a Currency chip selector (€/$/£/¥) to the Customise tab (on its own full-width row below output type/language, after Daniel flagged the first cramped 4-column layout) — pre-selected to the detected currency, user can override. `/api/generate` now accepts a `currency` form field and re-derives `mapped_values` with that override, so the generated One Pager/Slide Deck uses whichever currency the user picked. **Scope note:** the currency override only affects the generated PPT output (matches what Daniel asked for — "for the one pager or slides to use"); it does not trigger an exec-summary regenerate the way the language toggle does, since Claude's summary prompt doesn't typically quote exact budget figures.
- ✅ **Currency detection follow-up fix (same day)**: Daniel's first test showed the Budget tile as `$25,856.89` despite the source file clearly being in euros. Root cause: `detect_currency_symbol()` only catches currency symbols when the cell is literal *text* containing the character (e.g. `"€45,000"`) — but EOCs normally enter money as real numbers with a currency **number format** applied in Excel (e.g. a cell formatted `"€"#,##0.00`), which pandas reads as a plain float with no symbol at all. Added `detect_currency_from_workbook()`, which re-opens the file directly with `openpyxl` and checks each cell's `number_format` string for a currency symbol — this is what actually determines the € Daniel sees in Excel. `build_mapped_values()` now takes an `eoc_path` param (threaded through from all 8 call sites) and falls back to this workbook-format check whenever the cheap text-based check finds nothing. Verified against a synthetic workbook with a currency-number-formatted cell (no literal symbol text) — correctly detects `€`.
- ✅ **Added SGD (S$) to the currency list** (€/$/£/¥/S$ now), same day. `CURRENCY_SYMBOLS` in `main.py` is order-sensitive — `"S$"` must be checked/stripped *before* the bare `"$"` it contains, otherwise Singapore dollar values would always be misdetected/mis-stripped as plain USD. Both `detect_currency_symbol()`/`detect_currency_from_workbook()` (detection) and `safe_number()` (stripping before float parsing) updated with the correct order; verified with a quick script that `S$45,000` detects as `S$` (not `$`) and strips to a clean `45000.0`. Customise-tab currency row is now 5 chips (`grid-cols-5`).
- ✅ **Editable Markets tile on the Review screen**: when the EOC has no geo/market data, the Markets tile now shows a small pencil icon next to "Not available" (`app.html`, `renderTilesTo()` — only the `markets` field, scoped deliberately to just this tile per Daniel's ask rather than building a generic "any missing tile is editable" system). Clicking it swaps the tile for an inline text input (`startMarketsEdit`/`saveMarketsEdit`); the typed value is stored on `eoc.marketsOverride` and immediately reflected in the tile. At generate time, `markets_override` is sent as a form field to `/api/generate`, which overwrites `mapped_values["CAMPAIGN_MARKETS"]` after `build_mapped_values()` runs — same pattern as the currency/language overrides — so it flows into the actual PPT output. **Deliberately scoped:** editing Markets does *not* auto-regenerate the Claude exec summary (matches how editing the summary text or currency already work) — if Daniel wants the written summary to mention the newly-added markets, he needs to hit Regenerate himself afterward.
- ✅ **Currency chip sizing (same day)**: Daniel flagged the 5 currency chips as too large/two-line. Currency chips share the `.output-chip` CSS class (border, hover, selected state) with the Output Type and Language chips, so shrinking that base rule would've shrunk those too — instead added a `.currency-chip` modifier class (combined selector `.output-chip.currency-chip`, higher specificity so it reliably wins) with smaller padding, and changed the chip content from two stacked lines (symbol, then label) to one line: `€ - EUR`. Output Type and Language chips are untouched.
- ✅ **Editable Markets tile polish (same day)**: moved the edit pencil from next to the value ("FR ✏️") to next to the "MARKETS" label itself, per Daniel's preference. Also: editing Markets now updates the "See Full Validation Report" section too, not just the tile — that report renders from the raw markdown text (`display_validation_text`) the backend sent at validate time, which `renderFormattedReport()` parses into HTML but never re-fetches, so a client-side regex patches the `**Markets:**` line in that markdown to the override value before it's re-rendered (`renderReviewCard()`, right before the `renderFormattedReport()` call).

---

## Outstanding snag list

- [ ] Confirm the Round 7 changes (root redirect, legacy-domain redirect middleware) behave correctly once pushed and deployed — verify `https://pca.mobkoi.com/` lands on `/web`, and `https://pca-modular-builder-v12.onrender.com/anything` 307-redirects to the `pca.mobkoi.com` equivalent path.
- [ ] Confirm the JA wiring fix above actually produces a Japanese-templated deck + Japanese-toned summary when tested live (was fixed based on code inspection, not yet verified against a real generate call).
- [ ] Get Daniel's preferred Japanese phrase for "MRC Viewability" so the JA Slide Deck template's fallback label can be updated to match the English one.
- [ ] All of Round 6 (decimal fix, MRC rename, JA wiring fix, currency) needs a real local/staging test before pushing to `main` — none of it has been run against the live app yet, only verified via isolated logic tests and syntax checks (couldn't run the full FastAPI app in this sandbox — its SQLite init fails outside Daniel's actual environment).
- [x] Round 9's title-wrap fix — signed off 8 Sep 2026. Daniel tested live and caught a real bug (v1 false-triggered on ordinary one-line titles, e.g. "Rolex The Oscars UK 2026" got an unnecessary gap); fixed in the same session (see Round 9 below) and confirmed against his exact test case, plus a real generate-through-the-app spot check on both templates.

**Full Phase 1 feedback doc (read 23 July 2026 via shared Google Doc link, "PCA Generator - Feedback")** — collects notes from every regional office. Items marked "Actioned" in the doc that match this session's Round 6 work: Lesan/Yuri/Zoe's currency feedback (default-to-campaign-currency, JPY, SGD/USD), Yuri's OnScreen/MRC Viewability naming, Zoe's "AV delivery shows 1.54% instead of 154%". One Actioned item in the doc that we have **not** touched this session and should verify: Yuri's "Front Page: the layout looks broken" — worth confirming whether this was fixed separately or the doc is stale on this one.

Genuinely open items from the doc, not yet actioned:

- [ ] **Zoe (APAC SG) — Phase 2, noted in doc:** auto-detect market as "SG" from the campaign name; add a field for rate/deliverables (e.g. "CPX SGD 1.50, 150,000 Exposures"); text-input to populate the Campaign Snapshot slide (Overview/Audience/KPIs) or accept screen recordings.
- [ ] **Aiko (APAC JP) — Phase 2, noted in doc:** ability to update market-specific benchmarks, particularly for single-market campaigns.
- [ ] **Gaby (US):** choose which key metrics are shown (prioritise CTR/ER/VCR/Dwell Time, de-emphasise On-Screen); Dwell Time isn't available as a reportable metric at all; consolidate Campaign Details/Objectives/Key Figures into one slide; a dedicated slide per metric with a graph + benchmark line; a summary slide listing all sites with key metrics and highlighting above-benchmark performance; some text fields (e.g. "learnings") aren't consistently populating in English; simplify/skip full analysis section for US campaigns; option to show both Mobkoi and industry benchmarks side by side.
- [ ] **Erika (US):** metrics present in her report were reported by the tool as "not included" — needs a best-practice guide (or better fuzzy matching) for EOC column naming so the tool recognises them; general note that "the deck is bare bones."
- [ ] **Dorine (FR), not yet actioned:** bullet points for creative on the Campaign Analysis slide; standard boilerplate text on the Campaign Snapshot Overview ("Achieve awareness for the … and qualitative clicks to site in a premium editorial environment"); make Full Deck the default output selection (not One Pager); two additional slides — Publisher Overview and Creative Performance.

*(Awaiting Daniel's steer on which of the above to prioritise.)*

---

## Lessons learned

1. **Claude regression (Round 3):** the entire Claude integration was once silently removed from `web_routes.py` by an unrelated overwrite. Always verify a "working" feature with a screenshot before building on top of it. The `[claude] ANTHROPIC_API_KEY is not set` log line exists so this self-diagnoses in future.
1a. **Custom domain rollout (28 July 2026):** a second, broken Render service (`pca-automation-api`, Starter tier) existed alongside the real one (`pca-modular-builder`, Standard tier) — same GitHub repo/branch, but missing the `Root Directory: app` setting, so every deploy failed with `ModuleNotFoundError: No module named 'web_routes'`. It never served traffic and was safe to delete. **Lesson: if a Render URL 404s or a service looks unfamiliar, check the Service ID and Root Directory before assuming the working service is misconfigured — there may be an unrelated duplicate.** Separately: `PUBLIC_BASE_URL` feeding the OAuth `redirect_uri` means changing the canonical domain without updating this env var causes a cross-domain session-cookie mismatch (`mismatching_state: CSRF Warning!`) — this env var must be updated in lockstep with any domain change, alongside the Google Cloud Console authorized redirect URIs.
2. **Don't just defer on model choice when stakes are real** — show a concrete side-by-side before picking Haiku vs Sonnet, don't just recommend.
3. **Admin access is data-driven** (`ADMIN_EMAILS` env var), not DB-driven — intentional, keep it simple until there's a real need for more roles.
4. **Project snapshots/knowledge caches drift** — this file existing in the repo (rather than only in Claude project knowledge) is the fix. Treat the repo as ground truth; update this file after every session.
5. **Moving computers can flip file permission bits on every tracked file** (`644`→`755`), making git/GitHub Desktop show the whole repo as "changed" with no real content diff. Check `git diff --stat` (0 insertions/deletions = permissions only) before trusting a big change count; fix with `git config core.fileMode false` in the repo.
6. **A 30-day file-retention policy doesn't protect against a disk that fills up in less than 30 days** — Round 8's disk-full incident wasn't a cleanup bug, it was a disk sized too small for current volume. When diagnosing "is my cleanup/retention job working," check whether anything has actually aged past the retention window yet before assuming the job is broken.
7. **Double-check Render env vars actually match what the code assumes as a default** — `FILE_RETENTION_DAYS` was never set in Render (silently running on the code's default of 30) until Round 8. Don't assume a documented "default" env var is actually set in production.

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

## Round 7 (custom domain + domain migration, 28 July 2026)

- ✅ **Custom domain live:** `pca.mobkoi.com` added in Render (CNAME → `pca-modular-builder-v12.onrender.com`, added in AWS Route 53 since Mobkoi manages DNS there). Verified + SSL certificate issued.
- ✅ **Deleted the broken duplicate Render service** (`pca-automation-api`, Starter tier) — see Lessons learned below for root cause.
- ✅ **`PUBLIC_BASE_URL` updated** in Render env vars from the onrender.com URL to `https://pca.mobkoi.com` — fixes the OAuth `mismatching_state` CSRF error caused by the redirect_uri and session cookie being on mismatched domains after the domain switch.
- ✅ **Google Cloud Console OAuth client:** added `https://pca.mobkoi.com/auth/callback` as an additional authorized redirect URI, alongside the existing onrender.com one (kept as fallback).
- ✅ **Root path (`/`) now redirects to `/web`** instead of returning the raw health-check JSON — `main.py`'s `read_root()` changed to `return RedirectResponse(url="/web")`; `/health` is untouched and still returns the detailed status payload. `RedirectResponse` added to the `fastapi.responses` import.
- ✅ **Legacy-domain redirect middleware added** in `main.py`: any request arriving on `pca-modular-builder-v12.onrender.com` now gets a 307 redirect to the same path/query on `pca.mobkoi.com`, via a new `@app.middleware("http")` function (`redirect_legacy_onrender_domain`). This keeps old bookmarks/Slack links from breaking or hitting the cross-domain session-cookie issue described above, without needing to disable Render's "Render Subdomain" toggle. **Committed locally via Filesystem access + GitHub Desktop, not yet confirmed live** — check the deploy log and test both URLs after push.

---

## Round 8 (new machine move + disk-full incident, 12 Aug 2026)

- **Daniel moved to a new computer.** Local working copy re-cloned/copied to `pca-automation-api` on the new machine; remotes already pointed correctly at `Daniel-MOBKOI/pca-automation-api` (`origin`/`upstream`), old `danjcDesign` fork remote kept for reference but unused. Reconnected GitHub Desktop by adding the local folder and signing in with the account that has write access to the org repo.
- ✅ **False "49 changed files" in GitHub Desktop after the move — fixed, no real changes.** Every tracked file showed as modified purely because file permission bits flipped `644` → `755` during the transfer to the new machine (likely AirDrop/zip/external-drive copy). Content was byte-identical (`git diff --stat` showed `0 insertions(+), 0 deletions(-)` for every file; confirmed with a zip-level content diff on one `.pptx`). Fixed locally with `git config core.fileMode false` (run from inside the repo folder) — no commit needed, tells git to ignore permission-bit changes going forward.
- ✅ **Production incident: PCA generation failing with `[Errno 28] No space left on device`** — surfaced by a team member testing via email, saw "Generation failed" on the Done screen for a real PCA (Church's SS26). Root cause: the Render persistent disk was only 1GB (`/var/data`), and production usage filled it before the 30-day file-retention window ever had a chance to reclaim space.
  - **Cleanup logic itself verified working correctly, not buggy** — checked `cleanup_old_files()` in `app/web_routes.py` (background loop, runs every 24h) and confirmed via Render's Logs tab (`[cleanup] deleted=0 already_missing=1 errors=0 (retention=30d)` on both 7 & 8 Aug) that it correctly found nothing eligible to delete, because nothing in the DB was actually 30+ days old yet. The disk simply filled faster than the retention window could help — a capacity problem, not a code regression.
  - **Fix 1:** Render disk resized 1GB → 5GB. Confirmed (via web search of current Render docs) that resizing auto-triggers a Render deploy on save — no manual deploy step needed; new size is live once that deploy completes.
  - **Fix 2:** `FILE_RETENTION_DAYS` had never actually been set in Render's Environment tab (was silently relying on the code's `os.getenv(..., "30")` default). Added explicitly and set to `14` to rotate files off disk faster relative to disk size. Env var changes restart the service automatically, no code change/push required — the History page's "Files are kept for X days" text reads this same variable so it updated automatically too.
  - **How to verify the fix is holding, going forward:** Render → Logs tab, search `[cleanup]`, watch for `deleted=` counts starting to appear as files cross the new 14-day mark; or run `du -sh /var/data` / `find /var/data/generated -type f -mtime +14 | wc -l` in the Render Shell tab (Hobby-plan log retention is only 7 days, so the Shell/DB checks are the more durable way to check historically: `sqlite3 /var/data/pca.db "SELECT count(*) FROM runs WHERE created_at < datetime('now','-14 days') AND (one_pager_filename IS NOT NULL OR deck_filename IS NOT NULL);"` should return `0`).

---

## Round 9 (title-slide wrap/overlap fix, 8 Sep 2026 — signed off, not yet committed/pushed)

- **Bug (reported by Daniel, screenshot of the Slide Deck JA cover slide):** the title slide's `{{CAMPAIGN_NAME}}` text box was designed for one line, with the "PCA Reporting Results" / JA subtitle sitting right underneath with almost no gap. A campaign name long enough to wrap onto a second line visually covers the subtitle. Same tight-gap design exists on the One Pager's title block too (title → subtitle → exec summary).
- ✅ **Fixed in `app/main.py`** — no template file changes needed for the core fix, since it only affects runtime-generated output, not the editable masters. Added `find_title_slide_shapes()` (locates the CAMPAIGN_NAME title and any shapes that depend on its position — CAMPAIGN_PERIOD, EXEC_SUMMARY — via their `{{...}}` tokens, before substitution replaces them), `estimate_wrapped_line_count()` (line-wrap estimate), and `reflow_title_slide_for_wrapped_title()` (pushes the subtitle and dependent shapes down by the extra height, preserving their original relative spacing). Wired into both `build_grouped_stacked_modular_ppt` (One Pager) and `build_filtered_slide_deck_ppt` (Slide Deck) — covers both output types, EN and JA.
- **v1 → v2, same session:** the first version estimated line-wrap with a flat per-character multiplier and false-triggered on ordinary one-line titles — Daniel caught this live with a real campaign ("Rolex The Oscars UK 2026" got an unnecessary gap pushed under it, since it isn't actually two lines). Root cause: the flat multiplier was tuned too wide. Fixed by switching to real proportional measurement — bundled `app/fonts/DejaVuSans-Bold.ttf` (free/permissive license) and use Pillow to measure actual glyph widths, with a correction factor (`INTER_WIDTH_CORRECTION = 0.85`) since the title's real typeface (Inter) runs more compact than DejaVu Sans Bold, plus the text box's real internal margins rather than a guessed percentage. **New dependency:** `Pillow` added to `requirements.txt` — needs to install cleanly on the next Render deploy.
  - **Known limitation:** the real Inter font isn't available to measure directly in the generation environment (it's embedded/obfuscated inside the .pptx masters themselves, not extractable without implementing OOXML font de-obfuscation). `INTER_WIDTH_CORRECTION` is a calibrated approximation, not exact. If a title still visibly wraps unexpectedly, or triggers when it shouldn't, that's the constant to retune (see the "TITLE SLIDE REFLOW" comment block in `main.py`) — or drop a real Inter TTF/OTF into `app/fonts/` for exact measurement instead.
- **Verified:** re-tested "Rolex The Oscars UK 2026" specifically (correctly no longer shifts), plus a battery of short/medium/long titles across both templates and both languages (correctly shift only when genuinely wrapping), plus the One Pager's exec-summary-vs-next-content gap using a realistic 5-sentence Claude-style paragraph (stays clear, ~0.3in margin in the worst realistic case).
- **Separate incident found while testing, unrelated to this fix:** Daniel's translation edits to `modular_sections_master_ja.pptx` today (via whatever tool he used, before switching to native PowerPoint) flattened the TITLE_OVERVIEW slide's shape group — 52 shapes went from one group to individually loose, which breaks `build_grouped_stacked_modular_ppt` (it copies that section as a single group; ungrouped, it silently picked the largest loose shape — a picture — and produced near-blank One Pager output with no error). **Fixed by Daniel re-grouping natively in PowerPoint and re-saving** (confirmed 8 Sep 2026: slide 1 is back to one group of 52 children, and a full JA One Pager generates correctly again). **Lesson: a group flattening on save is a silent failure mode** — `build_grouped_stacked_modular_ppt` has no check that the TITLE_OVERVIEW section is actually still a group before copying it; worth a defensive check + clear error message another day, rather than relying on someone noticing blank output.
- Also surfaced: a few JA strings ("PCA Reporting Results" subtitle, "Campaign Objectives" heading) are intentionally left in English per Daniel/the Japan team's preference — not a bug, noted here so a future session doesn't "fix" it.
- **Status: signed off by Daniel 8 Sep 2026.** Code + font + requirements.txt change sitting locally, uncommitted (Daniel's standing rule — commits/pushes are his own step).

---

## Latest verified state

- **Commit:** `41c8382` (4 June 2026) — **Round 7 changes above are committed locally but the exact new commit hash is not yet known; update this once pushed.**
- **`app/web_routes.py`:** v10
- **`app/web_templates/base.html`:** v6
- **`app/web_templates/history.html`:** v8
- **`app/web_templates/admin.html`:** v2
- **`app/web_templates/landing.html`:** v3
- **`app/web_templates/denied.html`:** v2
- **`app/web_templates/app.html`:** language selector + word-by-word typewriter shipped; no version comment in file to confirm exact version number
- **`app/main.py`:** `template_path_for()` JA helper shipped; Round 9 added the title-slide reflow fix (`find_title_slide_shapes`, `estimate_wrapped_line_count`, `reflow_title_slide_for_wrapped_title`), wired into both PPT builders, signed off by Daniel — not yet committed/pushed
- **`app/fonts/DejaVuSans-Bold.ttf`:** new, bundled for Round 9's text-width measurement. **`requirements.txt`:** added `Pillow` for the same reason.
- **`app/templates/modular_sections_master_ja.pptx`:** Round 9 also fixed (by Daniel, in PowerPoint) a flattened shape group on slide 1 that broke JA One Pager generation — see Round 9 notes above.
