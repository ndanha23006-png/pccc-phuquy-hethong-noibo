# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Internal web app for **Công ty TNHH XD TM Phú Quý** (PCCC Phú Quý, a fire-safety construction company). It is a **Google Apps Script (GAS)** web app — no build tooling, no package manager, no tests. All source is plain HTML/CSS/JS and a single server-side `.gs` file, meant to be pasted directly into the Apps Script editor at https://script.google.com.

The repo has no local runtime: GAS-specific globals (`HtmlService`, `DriveApp`, `DocumentApp`, `SpreadsheetApp`, `MailApp`, `google.script.run`, etc.) only exist when the code is deployed inside an Apps Script project. There is nothing to `npm install` or `npm run` — "testing" a change means pasting it into the Apps Script editor (or its bound clasp project, if one is set up) and running/deploying it there.

## Repository layout

```
appscript/
  Code.gs               server-side backend: router (doGet) + all business logic
  appsscript.json        GAS manifest (scopes, webapp exec/access mode, timezone)
  styles.html             shared CSS, injected into every page via include('styles')
  index.html              home page — grid of tool cards linking to ?page=...
  phanhoi.html            "Xử lý phản hồi tổ đội" form (team feedback intake)
  baocaohub.html          weekly report hub — lists/creates week files
  baocao.html             weekly report entry form (per-team rows)
  HUONG_DAN_TRIEN_KHAI.md deployment guide (Vietnamese) — read this for setup/deploy steps
bieu_mau_xu_ly_phan_hoi_to_doi.html
                          standalone/offline prototype of the phanhoi form (pure client-side,
                          no google.script.run — computes a text record + mailto: link instead
                          of writing to Drive). Not wired into the GAS app; kept as a fallback/
                          demo version. Don't assume changes here propagate to appscript/phanhoi.html.
```

## Architecture

**Multi-page GAS web app, one feature per HTML file**, routed by a single `doGet(e)` in `Code.gs` via a `?page=` query param:

- `doGet` reads `e.parameter.page`, validates it against the `TRANG_HOP_LE` allow-list in `Code.gs`, defaults to `index` otherwise, and renders that template with `HtmlService.createTemplateFromFile(page)`.
- Every page template gets `baseUrl` (the deployed web app URL, for building `?page=...` links between pages) and `fileId` (used by `baocao.html` to know which week-file it's editing).
- Shared CSS lives once in `styles.html` and is pulled into every page's `<head>` via `<?!= include('styles') ?>` — `include()` in `Code.gs` just returns another HTML file's raw content, so it can also be used to pull in any other shared partial.
- Client-side JS in each page calls server functions via `google.script.run.withSuccessHandler(...).withFailureHandler(...)[fnName](...)`. `baocaohub.html`/`baocao.html` wrap this in a small `gs(fnName)` promise helper and also define a `MOCK` fallback so the page renders (with fake data) when opened outside GAS (`IS_GAS` check), for local preview in a plain browser.

**Two independent feature domains in `Code.gs`:**

1. **Phản hồi tổ đội (team feedback intake)** — `luuPhanHoi(d)`:
   - Server-side re-validates all required fields (never trust the client) — returns `{ok:false, thieu:[...]}` listing missing field labels if anything is empty.
   - Generates a ticket code `PH-yyyyMMdd-HHmmss`, builds a `DocumentApp` doc, appends a title/table of fields, exports it to `.docx` via the `.../export?...&exportFormat=docx` endpoint (needs `ScriptApp.getOAuthToken()`), saves it into `FOLDER_ID` on Drive, then trashes the temporary Google Doc.
   - Optionally emails the `.docx` to `EMAIL_NHAN_MAC_DINH` (or a per-submission override) via `MailApp` when `GUI_EMAIL` is true.

2. **Báo cáo công việc tuần (weekly work report)** — `listWeeks` / `createWeek` / `getWeekInfo` / `listTeams` / `addReport`:
   - One Google Sheets **spreadsheet per week** lives in `BAOCAO_FOLDER_ID`; `createWeek` creates it and moves it out of the root Drive folder into `BAOCAO_FOLDER_ID`.
   - Within a week's spreadsheet, **one sheet (tab) per team**, created lazily by `_teamSheet()` the first time that team reports (2 frozen header rows + `BC_HEADER` columns).
   - `addReport` takes `LockService.getScriptLock()` before writing (concurrent submissions from different teams/devices are expected), re-validates required row fields server-side, and appends rows with running row numbers, submitter email (`Session.getActiveUser().getEmail()`), and timestamp.
   - Sheet names named `HƯỚNG DẪN` or starting with `_` are treated as meta/instruction sheets and excluded from team listings via `_isMetaSheet()`.

**Config block at the top of `Code.gs`** (`FOLDER_ID`, `EMAIL_NHAN_MAC_DINH`, `GUI_EMAIL`, `BAOCAO_FOLDER_ID`, `TRANG_HOP_LE`) is the only place environment-specific values live — read `HUONG_DAN_TRIEN_KHAI.md` before changing these, since they're tied to specific Drive folder IDs that must exist and be shared correctly.

## Adding a new page/feature

Documented in `appscript/HUONG_DAN_TRIEN_KHAI.md` (section "THÊM TÍNH NĂNG MỚI SAU NÀY"):

1. Add a new `<name>.html` file with `<?!= include('styles') ?>` in its `<head>`.
2. Add `'<name>'` to `TRANG_HOP_LE` in `Code.gs`.
3. Link to it from `index.html` (or another page) via `href="<?= baseUrl ?>?page=<name>"`.
4. Add a handler function in `Code.gs`, called from the page via `google.script.run`.
5. In the actual Apps Script project, redeploy: **Deploy → Manage deployments → Edit → Version: New version → Deploy** (editing files alone does not update the live web app URL).

## Language

Always respond to the user in Vietnamese (tiếng Việt), regardless of the language used in code, comments, or file contents.

## Conventions

- All server-facing form functions **re-validate required fields on the server**, independent of client-side validation, and return a structured `{ok:false, ...}` rather than throwing, so the client can render field-level errors.
- UI copy, comments, and field labels are in **Vietnamese**; keep new user-facing strings consistent with that.
- Dates/times from `<input type="datetime-local">` are formatted with a manual regex (`fmt()` in `Code.gs`) rather than reparsed as JS `Date`, specifically to avoid timezone shifting the entered local time.
- Styling is a shared neumorphic design system (CSS custom properties `--bg`, `--card`, `--fire`, etc., `.neu`/`.neu-in` shadow utilities) defined once in `styles.html` — reuse those classes/variables rather than introducing new visual patterns.
