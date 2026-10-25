# ThesisSifu Pro — audit engine

FastAPI service (`main.py`) that audits a thesis (PDF/DOCX) and returns a ZIP of four reports.
Hosted on Railway (`web-production-88bb4.up.railway.app`); push to `main` → Railway redeploys.
The website (repo `thesis-aid-hub`, `src/pages/ThesisSifuPro.tsx`) calls **`POST /audit`** (not the
stream endpoint) and waits for the ZIP with no client-side timeout.

## Pipeline (v4.2)

| Stage | What | Model |
|---|---|---|
| 0 | Language detection from the thesis **body** (abstracts + reference list excluded) → `en` / `ms` | rules |
| 1 | Chapter + subsection split (English and BM headings) | regex |
| 2 | Spine extraction, document classification, **chapter digests** (what each chapter commits to, with quotes) — in parallel | Gemini 2.5 Flash |
| 3 | Subsection audit: 0–5 genuine issues each, with golden-thread link and literature search keywords | Gemini 2.5 Flash |
| 4 | Alignment audit (RQ trace, chapter-to-chapter links, construct consistency) ∥ literature grounding | Claude Sonnet 5 ∥ OpenAlex/Crossref + Gemini |
| 5 | Examiner report — runs **after** stage 4 and receives its results + verified literature | Claude Sonnet 5 |

Stages 4 and 5 receive the **full thesis text** (chapters, front matter excluded, prompt-cached) and are told to
verify every assistant-prepared point against it. The spine is automatic and may be incomplete: prompts say
"NOT FOUND" means extraction missed it, never that the thesis lacks it.

Outputs: `1_Examiner_Audit_Report.pdf`, `2_Annotated_Thesis.(docx|pdf)`, `3_Commentary_Report.pdf`,
`4_Alignment_Matrix_Report.pdf`. For BM theses every report, label and comment is in formal BM (DBP).

**Literature rule:** the model never writes a reference. Issues are searched in OpenAlex/Crossref and the
model may only pick from records actually returned (validated by id), so every cited work exists.

## Environment variables (Railway → Variables)

| Variable | Required | Purpose |
|---|---|---|
| `GEMINI_API_KEY` | yes | Gemini stages |
| `ANTHROPIC_API_KEY` | yes | Alignment + examiner |
| `OPENALEX_API_KEY` | optional | Better literature (abstracts, relevance). Free key = $1/day ≈ 1,000 searches; up to 24 per audit. Without it, Crossref is used. |
| `LITERATURE_MAILTO` | optional | Your email; puts Crossref requests in its faster "polite" pool |
| `CLAUDE_EFFORT` | optional | `high` (default) / `medium` / `low` — lower is faster and cheaper |
| `MAX_LIT_COMMENTS` | optional | How many issues get literature (default 24) |

Never commit keys; `.env` is gitignored.

## Endpoints
- `POST /audit` — multipart `file` → ZIP (used by the website)
- `POST /audit/stream` — SSE progress events (`progress`, `error`, `done`) + `: keepalive` comments; then
  `GET /audit/download/{audit_id}` (kept 10 min)
- `GET /health` — status, models, literature source

## Run locally
```
uv venv --python 3.12 .venv && uv pip install --python .venv/bin/python -r requirements.txt
GEMINI_API_KEY=... ANTHROPIC_API_KEY=... .venv/bin/uvicorn main:app --reload --port 8000
```

## Before merging to `main`
1. Deploy the branch to a Railway preview/staging service (not production).
2. Audit one **BM** and one **English** thesis with real keys; read the four reports.
3. Check total audit time — the site waits on one request, and stage 5 now runs after stage 4.
   If it gets long, set `CLAUDE_EFFORT=medium`.

## Status (2026-10-23)
v4.2 is on branch `improve/golden-thread-bm-literature`. Tested locally on real theses with **stubbed AI
responses** only: language detection, chapter splitting, BM RQ/RO extraction, live Crossref/OpenAlex search,
report building (BM + EN), SSE keep-alives, `/audit`, `/health`. **Not yet tested with real Gemini/Claude
calls** — feedback quality and audit duration are unverified.

## Why lecturers found v4.1 unhelpful (from real v4.1 reports, Oct 2026)
- Both Claude stages crashed in production (`temperature` rejected by anthropic SDK 2.x): no examiner report,
  empty alignment matrix, and the verdict banner read "FAIL" because it matched the word "failed".
- Spine said RQs/ROs "NOT FOUND" although the thesis had them; comments then told the student, as CRITICAL,
  that their RQs were missing.
- Chapter split broke (SPSS output read as chapter headings; unheaded Chapter 4 merged into Chapter 3), so
  results were judged as methodology.
- 196 comments with 91 CRITICAL; comments spoke the tool's jargon ("THESIS SPINE", "expected purpose");
  literature advice was generic; `YEAR_PLACEHOLDER` leaked into feedback.
- The examiner never read the thesis — it judged chapters by how many comments they got.

## Gotchas
- `anthropic` SDK 2.x rejects `temperature`; keep the pin `anthropic>=2.1,<3`.
- Gemini JSON is not schema-enforced — coerce types (`_flatten_to_str` / `_flatten_to_list`) at every boundary.
- Chapter headings in Word are often tab-indented (`\tBAB 2`) and chapter mentions appear in sentences
  ("Bab 1 ini membincangkan…"); the splitter handles both — keep the tests for them in mind when editing `_CH_PAT`.
