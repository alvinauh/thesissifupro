"""
ThesisSifu Pro v4.1 — Multi-Agent Thesis Panel (Alignment-Aware Edition)
========================================================================
FIXES in v4.1:
  - Migrated to google.genai SDK (google-generativeai deprecated)
  - Robust chapter splitter: handles "CHAPTER 1", "Chapter One", "1.", Roman numerals,
    ALL-CAPS titles, titles on next line, and plain numbered headings
  - Robust subsection splitter: no longer filters by chapter_num (breaks on Roman numerals);
    instead detects subsections by position within the chapter text slice
  - Safe Gemini response handling: checks candidates/finish_reason before .text access,
    with detailed error logging per call site
  - Page estimation fixed: searches actual chapter text for the excerpt,
    not just the subsection slice
  - Prompt injection guard: CANONICAL constants passed via .replace() not .format()
    to prevent KeyError on any curly braces in canonical text
  - Annotated PDF: bumped search window and added full-text fallback search
  - All exceptions logged with chapter+subsection context for Railway log debugging
  - v4.1.1: ThesisSpine fields are now coerced to their declared types (str/list)
    when built from Gemini's JSON. Gemini's output isn't schema-enforced, so a
    field documented as a string (e.g. "instrument") can come back as a JSON
    array, which used to crash every downstream .replace() call on that field
    with "TypeError: replace() argument 2 must be str, not list". Fixed by
    normalising every field once, at construction time, instead of trusting
    the JSON shape at every call site.

CHANGES in v4.2:
  - Language detection: the thesis BODY (abstracts and reference list excluded)
    decides English vs Bahasa Melayu; every stage and every report label then
    uses that language. BM headings (BAB, Persoalan/Objektif Kajian, Kerangka
    Teori, ...) are recognised by the splitter, purpose map and spine fallback.
  - Golden thread: each chapter is digested in full (Gemini) into what it
    commits to (RQs, constructs, measures, analyses, findings, contributions)
    with verbatim evidence. The alignment audit reads all digests, so it can
    trace a construct/RQ from Chapter 1 through to the conclusion.
  - The examiner now runs AFTER the alignment audit and receives its results
    (it previously ran in parallel and was handed an empty matrix).
  - Supporting literature: serious issues are matched to REAL published works
    retrieved from OpenAlex (OPENALEX_API_KEY) or Crossref; the model may only
    pick from retrieved records, so references cannot be invented.
  - Alignment + examiner use Claude Sonnet 5 with structured JSON output.
  - SSE keep-alives during long stages; /health route restored.

Pipeline:
  Stage 0  Language detection      (rules)    → en | ms
  Stage 1  Chapter + Subsection Split (regex) → structural map
  Stage 2  Spine + Classification + Chapter digests (Gemini, parallel)
  Stage 3  Subsection paragraph audit (Gemini, parallel)
  Stage 4  Alignment Audit (Claude) ∥ Literature grounding (OpenAlex/Crossref + Gemini)
  Stage 5  Holistic Examiner Report (Claude)  → uses spine + matrix + literature

Outputs (ZIP):
  1_Examiner_Audit_Report.pdf     — holistic critique, references spine
  2_Annotated_Thesis.(docx|pdf)   — subsection-aware inline comments
  3_Commentary_Report.pdf         — grouped by chapter → subsection → paragraph
  4_Alignment_Matrix_Report.pdf   — RQ↔RO↔Method↔Analysis↔Finding↔Conclusion

Endpoint:  POST /audit   multipart/form-data { file: <pdf|docx> }
Returns:   application/zip
"""

from __future__ import annotations

import io, os, re, json, asyncio, zipfile, hashlib, tempfile, traceback, base64
import urllib.request, urllib.parse
from datetime import datetime
from dataclasses import dataclass, field as dc_field
from typing import Optional, AsyncGenerator

import pypdf
import pdfplumber
import fitz  # PyMuPDF
from docx import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.opc.part import XmlPart
from docx.opc.packuri import PackURI

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from starlette.background import BackgroundTask

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY, TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, HRFlowable,
    Table, TableStyle, KeepTogether, PageBreak,
)

try:
    from google import genai as genai_sdk
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False


# ── App ────────────────────────────────────────────────────────
app = FastAPI(title="ThesisSifu Pro v4.2 — Alignment-Aware", version="4.2.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"],
                   allow_credentials=False, allow_methods=["*"], allow_headers=["*"])


# ── AI Clients ─────────────────────────────────────────────────
gemini_client = None
claude_client = None
GEMINI_MODEL  = "gemini-2.5-flash"   # current stable model as of June 2026
CLAUDE_MODEL  = "claude-sonnet-5"    # alignment + examiner (cross-chapter reasoning)
CLAUDE_EFFORT = os.environ.get("CLAUDE_EFFORT", "high")

# Literature grounding. OpenAlex is used when OPENALEX_API_KEY is set (free key =
# $1/day ≈ 1,000 searches); otherwise Crossref (free, no key). LITERATURE_MAILTO
# puts Crossref requests in its faster "polite" pool.
OPENALEX_API_KEY  = os.environ.get("OPENALEX_API_KEY", "")
LITERATURE_MAILTO = os.environ.get("LITERATURE_MAILTO", "")
MAX_LIT_COMMENTS  = int(os.environ.get("MAX_LIT_COMMENTS", 24))

if GEMINI_AVAILABLE:
    gkey = os.environ.get("GEMINI_API_KEY")
    if gkey:
        gemini_client = genai_sdk.Client(api_key=gkey)
        print(f"[ThesisSifu] Gemini {GEMINI_MODEL} ready (google.genai SDK)")
    else:
        print("[ThesisSifu] WARNING: GEMINI_API_KEY not set")

if ANTHROPIC_AVAILABLE:
    akey = os.environ.get("ANTHROPIC_API_KEY")
    if akey:
        claude_client = anthropic.AsyncAnthropic(api_key=akey)
        print(f"[ThesisSifu] Claude {CLAUDE_MODEL} ready")
    else:
        print("[ThesisSifu] WARNING: ANTHROPIC_API_KEY not set")


# ── Colours ────────────────────────────────────────────────────
W, H   = A4
BLACK  = colors.HexColor("#111111")
NAVY   = colors.HexColor("#0C2340")
ACCENT = colors.HexColor("#1A4A7A")
RULE   = colors.HexColor("#BBBBBB")
BOX_BG = colors.HexColor("#F4F6F9")
RED    = colors.HexColor("#A30000")
AMBER  = colors.HexColor("#7A4500")
GREEN  = colors.HexColor("#0A5C2A")
WHITE  = colors.white
LGRAY  = colors.HexColor("#EEEEEE")

SEV_COLORS = {"CRITICAL": RED, "MODERATE": AMBER, "SUGGESTION": ACCENT}
ALIGN_COLORS = {
    "ALIGNED":    GREEN,
    "PARTIAL":    AMBER,
    "MISALIGNED": RED,
    "MISSING":    colors.HexColor("#555555"),
    "UNCLEAR":    ACCENT,
}

TYPE_LABELS = {
    "JOURNAL_ARTICLE": "Scopus / Peer-Reviewed Journal Article",
    "UNDERGRADUATE":   "Undergraduate Thesis / Final Year Project",
    "MASTERS":         "Master's Thesis (By Research)",
    "PHD":             "PhD Dissertation",
}
TYPE_COLORS = {
    "JOURNAL_ARTICLE": colors.HexColor("#0A3D6B"),
    "UNDERGRADUATE":   colors.HexColor("#3A5A1C"),
    "MASTERS":         colors.HexColor("#6B3A00"),
    "PHD":             colors.HexColor("#4A0A0A"),
}


# ── Canonical libraries (NO curly braces — passed as plain text) ───────────
CANONICAL_FRAMEWORKS = (
    "THEORETICAL FRAMEWORKS (recommend by exact name + seminal author):\n"
    "\nTechnology / Information Systems:\n"
    "  - Technology Acceptance Model (TAM) — Davis (1989)\n"
    "  - UTAUT / UTAUT2 — Venkatesh et al. (2003, 2012)\n"
    "  - Diffusion of Innovations — Rogers (2003)\n"
    "  - Task-Technology Fit — Goodhue & Thompson (1995)\n"
    "  - DeLone & McLean IS Success Model — DeLone & McLean (2003)\n"
    "\nBehavioural / Psychological:\n"
    "  - Theory of Planned Behaviour — Ajzen (1991)\n"
    "  - Theory of Reasoned Action — Fishbein & Ajzen (1975)\n"
    "  - Social Cognitive Theory — Bandura (1986)\n"
    "  - Self-Determination Theory — Deci & Ryan (1985, 2000)\n"
    "  - Health Belief Model — Rosenstock (1974)\n"
    "\nStrategic Management / Business:\n"
    "  - Resource-Based View — Barney (1991)\n"
    "  - Dynamic Capabilities — Teece, Pisano & Shuen (1997)\n"
    "  - Porter's Five Forces — Porter (1980)\n"
    "  - Stakeholder Theory — Freeman (1984)\n"
    "  - Institutional Theory — DiMaggio & Powell (1983)\n"
    "\nOrganisational Behaviour / HR:\n"
    "  - Job Demands-Resources Model — Bakker & Demerouti (2007)\n"
    "  - Social Exchange Theory — Blau (1964)\n"
    "  - Transformational Leadership — Bass (1985)\n"
    "  - Psychological Capital — Luthans et al. (2007)\n"
    "\nMarketing / Consumer Behaviour:\n"
    "  - Stimulus-Organism-Response (SOR) — Mehrabian & Russell (1974)\n"
    "  - SERVQUAL — Parasuraman, Zeithaml & Berry (1988)\n"
    "  - Customer-Based Brand Equity — Keller (1993)\n"
    "  - Expectation-Confirmation Theory — Oliver (1980)\n"
    "\nEducation / Learning:\n"
    "  - Constructivism — Piaget, Vygotsky\n"
    "  - Bloom's Taxonomy (revised) — Anderson & Krathwohl (2001)\n"
    "  - Community of Inquiry — Garrison, Anderson & Archer (2000)\n"
    "  - TPACK — Mishra & Koehler (2006)\n"
)

CANONICAL_METHODS = (
    "METHODOLOGICAL REFERENCES (recommend by exact name + seminal author):\n"
    "\nQuantitative analytical techniques:\n"
    "  - PLS-SEM — Hair, Hult, Ringle & Sarstedt (2017, 2022)\n"
    "  - CB-SEM — Byrne (2016), Kline (2015)\n"
    "  - Hierarchical regression / moderation — Aiken & West (1991), Hayes PROCESS (2018)\n"
    "  - Mediation — Baron & Kenny (1986), Preacher & Hayes (2008)\n"
    "  - ANOVA/MANOVA — Tabachnick & Fidell (2019)\n"
    "  - Factor Analysis (EFA/CFA) — Hair et al. (2019), Brown (2015)\n"
    "\nQualitative analytical techniques:\n"
    "  - Thematic Analysis — Braun & Clarke (2006, 2019)\n"
    "  - Reflexive Thematic Analysis — Braun & Clarke (2022)\n"
    "  - Grounded Theory — Charmaz (2014), Strauss & Corbin (1998)\n"
    "  - Gioia Methodology — Gioia, Corley & Hamilton (2013)\n"
    "  - IPA — Smith, Flowers & Larkin (2009)\n"
    "  - Case Study — Yin (2018), Eisenhardt (1989)\n"
    "\nSystematic review:\n"
    "  - PRISMA 2020 — Page et al. (2021)\n"
    "  - Scoping review — Arksey & O'Malley (2005)\n"
    "\nSampling / sample size:\n"
    "  - Cochran's formula — Cochran (1977)\n"
    "  - Krejcie & Morgan table — Krejcie & Morgan (1970)\n"
    "  - G*Power power analysis — Faul et al. (2009)\n"
    "  - Minimum sample for PLS-SEM — Hair et al. (2017)\n"
    "\nValidity & reliability:\n"
    "  - Cronbach's alpha — Cronbach (1951)\n"
    "  - Composite reliability, AVE — Hair et al. (2017), Fornell & Larcker (1981)\n"
    "  - HTMT — Henseler, Ringle & Sarstedt (2015)\n"
    "  - Common Method Bias — Podsakoff et al. (2003, 2012)\n"
    "  - Content validity index (CVI) — Lynn (1986)\n"
    "\nQualitative trustworthiness:\n"
    "  - Lincoln & Guba (1985) — credibility, transferability, dependability, confirmability\n"
)


# ── Data classes ───────────────────────────────────────────────
@dataclass
class ThesisSpine:
    title:              str = "UNKNOWN"
    problem_statement:  str = "NOT FOUND"
    research_gap:       str = "NOT FOUND"
    research_questions: list = dc_field(default_factory=list)
    research_objectives:list = dc_field(default_factory=list)
    hypotheses:         list = dc_field(default_factory=list)
    theory_used:        str = "NOT FOUND"
    variables:          list = dc_field(default_factory=list)
    methodology:        str = "NOT FOUND"
    sampling:           str = "NOT FOUND"
    instrument:         str = "NOT FOUND"
    analysis_technique: str = "NOT FOUND"
    key_findings:       list = dc_field(default_factory=list)
    conclusions:        list = dc_field(default_factory=list)
    discipline:         str = "UNKNOWN"

@dataclass
class ParagraphComment:
    chapter:             str
    subsection:          str
    subsection_title:    str
    para_index:          int
    page_estimate:       int
    para_excerpt:        str
    severity:            str
    issue:               str
    recommendation:      str
    literature_needed:   str
    theory_needed:       str
    suggested_framework: str = ""
    suggested_method:    str = ""
    golden_thread_link:  str = ""      # which cross-chapter link this issue affects
    search_query:        str = ""      # keywords used to retrieve supporting literature
    references:          list = dc_field(default_factory=list)  # verified works, see attach_literature

@dataclass
class Subsection:
    chapter_num:      str
    subsection_num:   str
    title:            str
    text:             str
    char_offset:      int = 0          # offset within full_text for page estimation
    expected_purpose: str = ""
    comments:         list = dc_field(default_factory=list)

@dataclass
class ChapterSummary:
    chapter_num:   str
    chapter_title: str
    text:          str = ""
    char_offset:   int = 0             # offset within full_text for page estimation
    subsections:   list = dc_field(default_factory=list)
    comments:      list = dc_field(default_factory=list)
    digest:        dict = dc_field(default_factory=dict)   # what this chapter commits to

@dataclass
class AlignmentRow:
    rq:               str = "—"
    ro:               str = "—"
    hypothesis:       str = "—"
    theory_construct: str = "—"
    method:           str = "—"
    instrument:       str = "—"
    analysis:         str = "—"
    finding:          str = "—"
    discussion:       str = "—"
    conclusion:       str = "—"
    status:           str = "UNCLEAR"
    broken_link:      str = ""
    evidence:         str = ""
    note:             str = ""
    fix:              str = ""

@dataclass
class AlignmentMatrix:
    rows:                       list = dc_field(default_factory=list)
    overall_verdict:            str = ""
    golden_thread_score:        str = "UNCLEAR"
    critical_gaps:              list = dc_field(default_factory=list)
    structural_recommendations: list = dc_field(default_factory=list)
    chapter_links:              list = dc_field(default_factory=list)  # dicts: from_chapter, to_chapter, link, status, evidence, fix
    construct_consistency:      list = dc_field(default_factory=list)  # dicts: construct, trace, status, note


# ── Gemini response helper ─────────────────────────────────────
def _gemini_text(response, context: str = "") -> str:
    """Safely extract text from a google.genai response.
    Logs the finish reason if the response was blocked or empty."""
    try:
        if not response or not response.candidates:
            print(f"[ThesisSifu] Gemini: no candidates returned. Context: {context}")
            return ""
        cand = response.candidates[0]
        # finish_reason: 1=STOP 2=MAX_TOKENS 3=SAFETY 4=RECITATION 5=OTHER
        reason = getattr(cand, "finish_reason", None)
        if reason and reason not in (1, "STOP"):
            print(f"[ThesisSifu] Gemini blocked/incomplete. reason={reason} context={context}")
            return ""
        return response.text or ""
    except Exception as e:
        print(f"[ThesisSifu] _gemini_text error ({context}): {e}")
        return ""


# ── Spine field type-coercion helpers ───────────────────────────
# Gemini's JSON output is NOT schema-enforced. A field the prompt describes
# as a plain string (e.g. "instrument": "questionnaire / interview protocol")
# can legally come back as a JSON array, dict, or number instead. If that
# value is assigned straight into a `str`-typed ThesisSpine field, every
# downstream `.replace(PLACEHOLDER, spine.some_field)` call in
# run_alignment_audit / run_examiner / audit_subsection crashes with
# "TypeError: replace() argument 2 must be str, not list" — as happened in
# production. Coercing once here, at construction time, is cheaper and safer
# than guarding every call site individually.
def _flatten_to_str(v, sep: str = "; ") -> str:
    """Coerce any Gemini JSON value into plain text for a str-typed spine field."""
    if v is None:
        return ""
    if isinstance(v, (list, tuple, set)):
        return sep.join(_flatten_to_str(x, sep) for x in v)
    if isinstance(v, dict):
        return sep.join(f"{k}: {_flatten_to_str(val, sep)}" for k, val in v.items())
    return str(v)

def _flatten_to_list(v) -> list:
    """Coerce any Gemini JSON value into a list for a list-typed spine field.
    Guards the mirror-image failure mode: if Gemini returns a bare string where
    a list was expected, code like '; '.join(spine.research_questions) would
    silently join individual characters instead of crashing — a quieter but
    still-wrong outcome that this normalises away."""
    if v is None:
        return []
    if isinstance(v, list):
        return v
    if isinstance(v, (tuple, set)):
        return list(v)
    return [str(v)]


def _parse_json_loose(raw: str):
    """Parse model JSON output, tolerating markdown fences."""
    raw = re.sub(r"```(?:json)?", "", raw or "").strip("`").strip()
    return json.loads(raw)


async def _gemini_json(prompt: str, context: str):
    """One Gemini call in JSON mode. Returns parsed JSON, or None on any failure."""
    if not gemini_client:
        return None
    raw = ""
    try:
        r = await gemini_client.aio.models.generate_content(
            model=GEMINI_MODEL, contents=prompt,
            config={"response_mime_type": "application/json"})
        raw = _gemini_text(r, context)
        if not raw:
            return None
        return _parse_json_loose(raw)
    except json.JSONDecodeError as e:
        print(f"[ThesisSifu] JSON parse error in {context}: {e}\nRaw: {raw[:300]}")
    except Exception as e:
        print(f"[ThesisSifu] Gemini error in {context}: {e}")
    return None


# ── Language detection ─────────────────────────────────────────
# Malaysian theses usually carry BOTH an English Abstract and a BM Abstrak, and
# the reference list is mostly English even in a BM thesis. Counting function
# words over the whole file therefore misjudges the language, so the abstracts
# and the reference list are removed first and only the body is counted.
_BM_WORDS = frozenset(
    "yang dan dengan untuk dalam ini itu kepada adalah oleh bagi tidak pada akan "
    "telah dapat daripada iaitu serta juga atau secara terhadap boleh lebih "
    "kajian antara mereka sebagai".split())
_EN_WORDS = frozenset(
    "the and of to in is that for with this are was be by which as on from not "
    "were their these have has been an it can between".split())

_ABSTRACT_HEAD = re.compile(r'^[ \t]*(?:ABSTRACT|ABSTRAK)\b', re.IGNORECASE | re.MULTILINE)
_ABSTRACT_END = re.compile(
    r'^[ \t]*(?:ABSTRACT|ABSTRAK|ACKNOWLEDGE?MENTS?|PENGHARGAAN|TABLE OF CONTENTS|'
    r'CONTENTS|KANDUNGAN|ISI KANDUNGAN|LIST OF TABLES|SENARAI JADUAL|DECLARATION|'
    r'PENGAKUAN|DEDICATION|DEDIKASI|CHAPTER\s+\S+|BAB\s+\S+)\b',
    re.IGNORECASE | re.MULTILINE)
_REFS_HEAD = re.compile(r'^[ \t]*(?:REFERENCES|RUJUKAN|BIBLIOGRAPHY|BIBLIOGRAFI)[ \t]*$',
                        re.IGNORECASE | re.MULTILINE)
# "Smith, J. (2020). ..." / "[12] Smith, J., ..." / "Ahmad Zaki, M. (2019)."
_REF_ENTRY = re.compile(r'^\s*(?:\[\d+\]\s*)?[A-Z][\w\'\-]+(?:\s[A-Z][\w\'\-]+)*,\s+[A-Z]\.[^\n]*\(?\d{4}')

def _strip_abstracts(text: str) -> str:
    out, pos = [], 0
    for m in _ABSTRACT_HEAD.finditer(text):
        if m.start() < pos:
            continue
        out.append(text[pos:m.start()])
        nxt = _ABSTRACT_END.search(text, m.end())
        # An abstract is at most a page or two; never drop more than ~8k chars.
        pos = nxt.start() if nxt and nxt.start() - m.end() < 8000 else min(len(text), m.end() + 4000)
    out.append(text[pos:])
    return "".join(out)

def detect_language(full_text: str) -> dict:
    """Decide the thesis language from its body (abstracts + references excluded)."""
    body = _strip_abstracts(full_text)
    # Cut the reference list + appendices at the last REFERENCES/RUJUKAN heading,
    # unless that would discard most of the text (then it is probably a ToC entry).
    refs = list(_REFS_HEAD.finditer(body))
    if refs and refs[-1].start() > len(body) * 0.4:
        body = body[:refs[-1].start()]
    # Reference entries can also sit at chapter ends; drop lines shaped like one.
    body = "\n".join(l for l in body.split("\n") if not _REF_ENTRY.match(l))
    words = re.findall(r"[a-z]+", body.lower())
    bm = sum(w in _BM_WORDS for w in words)
    en = sum(w in _EN_WORDS for w in words)
    ratio = bm / (bm + en) if bm + en else 0.0
    return {"lang": "ms" if ratio > 0.5 else "en", "bm_ratio": round(ratio, 2)}

def _lang_instruction(lang: str) -> str:
    if lang == "ms":
        return ("LANGUAGE: The thesis body is written in Bahasa Melayu. Write ALL feedback "
                "(every sentence you produce, including every JSON string value) in formal "
                "academic Bahasa Melayu following Dewan Bahasa dan Pustaka (DBP) spelling and "
                "terminology. Keep names of theories, models, methods, statistical tests, authors "
                "and citations in their original form (e.g. 'Theory of Planned Behaviour (Ajzen, "
                "1991)', 'PLS-SEM'). Quote the thesis verbatim in its original language. Keep JSON "
                "keys and the fixed labels (CRITICAL, MODERATE, SUGGESTION, ALIGNED, PARTIAL, "
                "MISALIGNED, MISSING, UNCLEAR, STRONG, ACCEPTABLE, WEAK, BROKEN) in English exactly "
                "as specified.")
    return ("LANGUAGE: Write all feedback in formal academic English. Quote the thesis "
            "verbatim in its original language.")

# Report labels: (English, Bahasa Melayu)
_LABELS = {
    "examiner_title":   ("EXAMINER'S AUDIT REPORT", "LAPORAN AUDIT PEMERIKSA"),
    "commentary_title": ("COMMENTARY REPORT", "LAPORAN ULASAN"),
    "commentary_badge": ("PARAGRAPH-LEVEL COMMENTARY REPORT", "LAPORAN ULASAN PERINGKAT PERENGGAN"),
    "commentary_sub":   ("Subsection-Level Supervisor Notes", "Catatan Penyelia Peringkat Subseksyen"),
    "align_title":      ("ALIGNMENT MATRIX REPORT", "LAPORAN MATRIKS PENJAJARAN"),
    "align_badge":      ("STRUCTURAL ALIGNMENT MATRIX", "MATRIKS PENJAJARAN STRUKTUR"),
    "align_sub":        ("Golden Thread Audit", "Audit Benang Emas (Golden Thread)"),
    "spine":            ("EXTRACTED THESIS SPINE", "TULANG BELAKANG TESIS YANG DIEKSTRAK"),
    "spine_ref":        ("THESIS SPINE (for reference)", "TULANG BELAKANG TESIS (rujukan)"),
    "golden_thread":    ("Golden Thread", "Benang Emas"),
    "align_score_sub":  ("Structural Alignment Score", "Skor Penjajaran Struktur"),
    "verdict_sub":      ("Examiner's Recommended Outcome", "Keputusan Disyorkan Pemeriksa"),
    "issue_summary":    ("CHAPTER + SUBSECTION ISSUE SUMMARY", "RINGKASAN ISU BAB + SUBSEKSYEN"),
    "overall_verdict":  ("OVERALL ALIGNMENT VERDICT", "RUMUSAN KESELURUHAN PENJAJARAN"),
    "rq_trace":         ("ALIGNMENT MATRIX — RQ-LEVEL TRACE", "MATRIKS PENJAJARAN — JEJAK PERSOALAN KAJIAN"),
    "row_notes":        ("ROW-BY-ROW NOTES", "CATATAN SETIAP BARIS"),
    "chapter_links":    ("CHAPTER-TO-CHAPTER LINKS", "PAUTAN ANTARA BAB"),
    "constructs":       ("CONSTRUCT CONSISTENCY ACROSS CHAPTERS", "KONSISTENSI KONSTRUK MERENTAS BAB"),
    "critical_gaps":    ("CRITICAL STRUCTURAL GAPS", "JURANG STRUKTUR KRITIKAL"),
    "struct_recs":      ("STRUCTURAL RECOMMENDATIONS", "CADANGAN STRUKTUR"),
    "scorecard":        ("CHAPTER DELIVERY SCORECARD", "KAD SKOR PENYAMPAIAN BAB"),
    "distribution":     ("OVERALL COMMENT DISTRIBUTION", "TABURAN KESELURUHAN ULASAN"),
    "issue":            ("Issue", "Isu"),
    "recommendation":   ("Recommendation", "Cadangan"),
    "literature":       ("Literature needed", "Literatur diperlukan"),
    "theory":           ("Theory / Framework", "Teori / Kerangka"),
    "framework":        ("Suggested framework", "Kerangka dicadangkan"),
    "method":           ("Suggested method/reference", "Kaedah/rujukan dicadangkan"),
    "thread_link":      ("Golden-thread link affected", "Pautan benang emas terjejas"),
    "verified_refs":    ("Supporting literature (verified records)", "Literatur sokongan (rekod disahkan)"),
    "expected":         ("Expected purpose", "Tujuan dijangka"),
    "evidence":         ("Evidence", "Bukti"),
    "fix":              ("Fix", "Pembetulan"),
    "broken_link":      ("Broken link", "Pautan terputus"),
}
_SEV_LABELS_MS = {"CRITICAL": "KRITIKAL", "MODERATE": "SEDERHANA", "SUGGESTION": "CADANGAN"}

def L(key: str, lang: str) -> str:
    en, ms = _LABELS[key]
    return ms if lang == "ms" else en

def sev_label(sev: str, lang: str) -> str:
    return _SEV_LABELS_MS.get(sev, sev) if lang == "ms" else sev


# ── Text extraction ─────────────────────────────────────────────
def extract_text(content: bytes, filename: str) -> str:
    """Extract text from PDF or DOCX.
    Uses pdfplumber for PDFs (preserves inter-word spaces that pypdf often drops),
    with pypdf as fallback. Word-concatenation like '1Introduction' or '3ProblemStatement'
    causes the chapter/subsection regex to produce hundreds of false positives,
    which then starves the Gemini subsection auditor of real text."""
    fn = filename.lower()
    if fn.endswith(".pdf"):
        # Primary: pdfplumber (layout-aware, preserves spaces)
        try:
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                pages = []
                for page in pdf.pages:
                    t = page.extract_text(x_tolerance=1, y_tolerance=2) or ""
                    pages.append(t)
                text = "\n".join(pages)
                if text.strip():
                    print(f"[ThesisSifu] pdfplumber extracted {len(text)} chars")
                    return text
        except Exception as e:
            print(f"[ThesisSifu] pdfplumber failed ({e}), falling back to pypdf")
        # Fallback: pypdf
        try:
            reader = pypdf.PdfReader(io.BytesIO(content))
            pages = []
            for page in reader.pages:
                t = page.extract_text() or ""
                pages.append(t)
            text = "\n".join(pages)
            print(f"[ThesisSifu] pypdf fallback extracted {len(text)} chars")
            return text
        except Exception as e:
            print(f"[ThesisSifu] pypdf also failed: {e}")
            return ""
    if fn.endswith(".docx"):
        doc = DocxDocument(io.BytesIO(content))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    return ""


# ── Chapter + Subsection splitter (robust) ────────────────────
# Chapter pattern: ONLY matches "CHAPTER N" / "CHAPTER I" / "CHAPTER ONE"
# and "N.0 Title" forms (explicit section-zero heading).
# The old (\d+)[:\.\-] branch was matching statistics (3.432, 0.521) as chapters.
_WORD_NUMS = {"one":"1","two":"2","three":"3","four":"4","five":"5",
              "six":"6","seven":"7","eight":"8","nine":"9","ten":"10",
              "satu":"1","dua":"2","tiga":"3","empat":"4","lima":"5",
              "enam":"6","tujuh":"7","lapan":"8","sembilan":"9","sepuluh":"10"}
_ROMAN = {"I":"1","II":"2","III":"3","IV":"4","V":"5",
          "VI":"6","VII":"7","VIII":"8","IX":"9","X":"10"}

_CH_PAT = re.compile(
    r'(?:^|\n)[ \t]*'          # Word often indents/centres headings with tabs ("\tBAB 2")
    r'(?:'
      # "CHAPTER 1", "CHAPTER I", "CHAPTER ONE", "BAB 1", "BAHAGIAN 2"
      r'(?:CHAPTER|BAHAGIAN|BAB)\s+(\d{1,2}|[IVX]+|ONE|TWO|THREE|FOUR|FIVE|SIX|SEVEN|EIGHT|NINE|TEN'
      r'|SATU|DUA|TIGA|EMPAT|LIMA|ENAM|TUJUH|LAPAN|SEMBILAN|SEPULUH)(?:\.0+)?\b'   # "CHAPTER 1.0 :"
      # "1.0 Introduction" or "2.0 Literature Review" — ONLY N.0 forms followed by a word,
      # never statistics or SPSS output ("1.00   8 . 0", "3.432")
      r'|(\d{1,2})\.0+[ \t]+(?=[A-Za-z])'
    r')'
    r'[:\s\-—–]*([^\n]{0,100})',
    re.IGNORECASE,
)

# Subsection: "1.1 Title", "3.3.1 Title" — integers only, MUST be followed by space+letter.
# Lookahead (?=[ \t]+[A-Za-z]) prevents matching statistics like 3.432, 0.521, 1.015.
_SUB_PAT = re.compile(
    r'(?:^|\n)[ \t]*'
    r'(\d{1,2}\.\d{1,3}(?:\.\d{1,3})?)'   # N.M or N.M.P (short integers, max 3 digits each)
    r'(?=[ \t]+[A-Za-z])'                   # LOOKAHEAD: must be followed by whitespace then a letter
    r'[ \t]+([^\n]{0,120})',                 # capture the title
)

_SUBSECTION_PURPOSE = {
    "background":           "Establish context and historical/practical relevance of the problem.",
    "problem statement":    "Articulate the specific research problem with empirical/theoretical evidence.",
    "research gap":         "Demonstrate what prior work has NOT addressed.",
    "research question":    "Pose 1+ specific, answerable research questions tied to the problem.",
    "objective":            "List measurable objectives that map 1-to-1 to research questions.",
    "hypothes":             "State testable, directional hypotheses linked to theory.",
    "significance":         "Explain theoretical and practical contributions of the study.",
    "scope":                "Define boundaries (population, geography, time, constructs).",
    "definition":           "Operationalise key terms used in the study.",
    "theoretical framework":"Justify the underpinning theory and link constructs to it.",
    "conceptual framework": "Present the model/diagram of relationships among variables.",
    "literature review":    "Synthesise prior research critically, not just summarise.",
    "research design":      "Justify quantitative/qualitative/mixed approach and paradigm.",
    "population":           "Define the target population and sampling frame.",
    "sampling":             "Justify sampling technique and sample size with a recognised formula.",
    "instrument":           "Describe instrument development, items, scale, and pretest.",
    "validity":             "Report content/construct validity evidence (e.g., CVI, AVE).",
    "reliability":          "Report reliability evidence (Cronbach's alpha, composite reliability).",
    "data collection":      "Detail procedure, ethics approval, response rate.",
    "data analysis":        "Justify analytical technique against research questions.",
    "demographic":          "Present sample profile relevant to research questions.",
    "descriptive":          "Report descriptives that contextualise constructs.",
    "results":              "Present findings against each research question/hypothesis explicitly.",
    "finding":              "Report findings tied directly to each RQ.",
    "discussion":           "Interpret findings against literature; explain unexpected results.",
    "implication":          "Theoretical and practical implications grounded in findings.",
    "limitation":           "Acknowledge methodological and scope limitations honestly.",
    "future research":      "Propose specific, actionable next studies.",
    "conclusion":           "Restate contribution and answer each research question explicitly.",
    "summary":              "Concise restatement of findings against objectives.",
    "recommendation":       "Specific, actionable recommendations for stakeholders.",
}

# Bahasa Melayu headings → the English key whose purpose applies.
# Order matters: longer / more specific phrases first.
_BM_PURPOSE_ALIASES = [
    ("latar belakang", "background"), ("pernyataan masalah", "problem statement"),
    ("penyataan masalah", "problem statement"), ("jurang kajian", "research gap"),
    ("persoalan kajian", "research question"), ("soalan kajian", "research question"),
    ("objektif", "objective"), ("hipotesis", "hypothes"),
    ("kepentingan kajian", "significance"), ("signifikan", "significance"),
    ("skop", "scope"), ("batasan kajian", "limitation"), ("limitasi", "limitation"),
    ("definisi", "definition"), ("kerangka teori", "theoretical framework"),
    ("kerangka teoretikal", "theoretical framework"), ("kerangka konsep", "conceptual framework"),
    ("sorotan literatur", "literature review"), ("tinjauan literatur", "literature review"),
    ("kajian lepas", "literature review"), ("reka bentuk", "research design"),
    ("populasi", "population"), ("persampelan", "sampling"), ("sampel", "sampling"),
    ("instrumen", "instrument"), ("kesahan", "validity"), ("kebolehpercayaan", "reliability"),
    ("pengumpulan data", "data collection"), ("prosedur", "data collection"),
    ("analisis data", "data analysis"), ("demografi", "demographic"),
    ("deskriptif", "descriptive"), ("dapatan", "finding"), ("keputusan", "results"),
    ("perbincangan", "discussion"), ("implikasi", "implication"),
    ("kajian lanjutan", "future research"), ("kesimpulan", "conclusion"),
    ("rumusan", "summary"), ("cadangan", "recommendation"),
]

def _purpose_for(title: str) -> str:
    t = title.lower()
    for k, v in _SUBSECTION_PURPOSE.items():
        if k in t:
            return v
    for bm, key in _BM_PURPOSE_ALIASES:
        if bm in t:
            return _SUBSECTION_PURPOSE[key]
    return "Auditor should infer purpose from chapter context."


# ── Chapter roles + what examiners check in each ──────────────
_ROLE_KEYWORDS = [
    ("literature_review", ("literature", "sorotan", "tinjauan", "kajian lepas", "kajian literatur")),
    ("methodology",       ("method", "metodologi", "kaedah", "reka bentuk", "research design")),
    ("discussion",        ("discussion", "perbincangan")),
    ("results",           ("result", "finding", "data analysis", "analisis data", "dapatan", "analysis")),
    ("conclusion",        ("conclusion", "kesimpulan", "rumusan", "recommendation", "cadangan", "summary")),
    ("introduction",      ("introduction", "pengenalan", "pendahuluan", "background", "latar belakang")),
]
_ROLE_BY_POSITION = {"1": "introduction", "2": "literature_review", "3": "methodology",
                     "4": "results", "5": "conclusion"}

_CHAPTER_CHECKS = {
    "introduction": (
        "- Problem is evidenced (statistics, reports, prior studies), local AND international.\n"
        "- Gap is specific (which population / context / construct / method is missing) and cited.\n"
        "- RQs and ROs map 1-to-1 with measurable verbs; hypotheses follow from the theory.\n"
        "- Scope and operational definitions match the constructs used in later chapters.\n"
        "- Significance names concrete theoretical AND practical contributions."),
    "literature_review": (
        "- Synthesis, not author-by-author summary: themes, agreement/disagreement, trends.\n"
        "- Every construct in the RQs is defined, reviewed and its measurement discussed.\n"
        "- The underpinning theory is explained and explicitly linked to each hypothesised relationship.\n"
        "- Currency: substantial share of sources from the last 5-7 years, plus seminal works.\n"
        "- The gap emerges from this review and matches the gap claimed in Chapter 1.\n"
        "- Conceptual framework is derived from theory + literature; every path is justified."),
    "methodology": (
        "- Paradigm and design are justified against the RQs.\n"
        "- Population, sampling frame, technique and sample size justified (formula / power analysis).\n"
        "- Instrument items map to EVERY construct; adapted items cite their source.\n"
        "- Validity and reliability evidence with accepted thresholds; pilot study reported.\n"
        "- Analysis technique stated per RQ/hypothesis; ethics and data procedures described."),
    "results": (
        "- Results are organised by RQ/hypothesis and every RQ/hypothesis is answered.\n"
        "- Data screening / assumptions reported; statistics with effect sizes and thresholds.\n"
        "- Tables and figures are interpreted, not just listed."),
    "discussion": (
        "- Each finding is interpreted against the Chapter 2 literature (agree/disagree and why).\n"
        "- The underpinning theory is used to explain the findings; unexpected results explained.\n"
        "- No new results introduced."),
    "conclusion": (
        "- Each RQ is answered explicitly.\n"
        "- Contributions match those promised in Chapter 1 and are supported by the findings.\n"
        "- Implications are grounded in findings; no overclaiming beyond sample and design.\n"
        "- Limitations are honest; future research is specific and actionable."),
    "other": "- Infer the section's role and judge whether it serves the thesis argument.",
}

def chapter_role(cs) -> str:
    role = (cs.digest or {}).get("chapter_role", "")
    if role in _CHAPTER_CHECKS:
        return role
    t = cs.chapter_title.lower()
    for r, kws in _ROLE_KEYWORDS:
        if any(k in t for k in kws):
            # Every chapter opens with an "Introduction"/"Pendahuluan" subsection, which
            # the splitter sometimes takes as the chapter title; only trust it for Ch 1.
            if r == "introduction" and cs.chapter_num != "1":
                break
            return r
    return _ROLE_BY_POSITION.get(cs.chapter_num, "other")

import textwrap

def _shorten(s: str, width: int) -> str:
    return textwrap.shorten(s or "", width=width, placeholder="…")

def _normalise_chnum(raw: str) -> str:
    """Convert Roman/word chapter numbers to Arabic integers."""
    r = raw.strip().upper()
    if r in _ROMAN: return _ROMAN[r]
    if r.lower() in _WORD_NUMS: return _WORD_NUMS[r.lower()]
    return raw.strip()

def split_chapters(text: str) -> list[ChapterSummary]:
    matches = list(_CH_PAT.finditer(text))
    if not matches:
        size = 6000
        return [ChapterSummary(chapter_num=str(i+1),
                               chapter_title=f"Section {i+1}",
                               text=text[s:s+size],
                               char_offset=s)
                for i, s in enumerate(range(0, len(text), size))]

    # Build candidate spans first, THEN dedupe — keep the largest span per
    # chapter number. A real chapter body is always much longer than a
    # one-line Table-of-Contents entry, so this drops ToC noise matches
    # and the garbled "title bleeds into next heading" problem they cause.
    candidates = []
    for i, m in enumerate(matches):
        raw_num = (m.group(1) or m.group(2) or str(i+1))
        num     = _normalise_chnum(raw_num)
        title   = (m.group(3) or "").strip() or f"Chapter {num}"
        # Stop the title at the first sign it has bled into a numbered
        # subsection heading (e.g. "LITERATURE REVIEW 2.1 Introduction")
        title = re.split(r'\s+\d{1,2}\.\d', title)[0].strip()
        # "Bab 1 ini membincangkan ..." / "Chapter 2 reviews ..." is a sentence
        # referring to a chapter, not a heading: headings aren't lower-case.
        if title and title[0].islower():
            continue
        if not title:
            next_nl = text.find("\n", m.end())
            title = text[m.end():next_nl].strip() if next_nl != -1 else f"Chapter {num}"
        start = m.start()
        end   = matches[i+1].start() if i+1 < len(matches) else len(text)
        candidates.append((num, title, start, end))

    best = {}
    for num, title, start, end in candidates:
        span = end - start
        if num not in best or span > (best[num][3] - best[num][2]):
            best[num] = (num, title, start, end)

    chapters = []
    for num, title, start, end in sorted(best.values(), key=lambda x: x[2]):
        chapters.append(ChapterSummary(
            chapter_num=num, chapter_title=title,
            text=text[start:end].strip(), char_offset=start))
    _recover_chapter_one(text, chapters)
    _split_unheaded_chapters(chapters)
    return chapters or [ChapterSummary("1", "Full Document", text, 0)]

def _split_unheaded_chapters(chapters: list) -> None:
    """Some theses omit a chapter heading (e.g. no "CHAPTER 4"), so its sections
    (4.1, 4.2 ...) end up inside the previous chapter and get judged against the
    wrong chapter's purpose. Split where numbering moves to a missing chapter."""
    i = 0
    while i < len(chapters):
        cs = chapters[i]
        have = {c.chapter_num for c in chapters}
        nxt = str(int(cs.chapter_num) + 1) if cs.chapter_num.isdigit() else None
        m = None
        if nxt and nxt not in have:
            m = re.search(r'(?:^|\n)[ \t]*' + nxt + r'\.1[ \t:]+[A-Za-z][^\n]*', cs.text)
        if m and m.start() > 1000:
            body = cs.text[m.start():].strip()
            cs.text = cs.text[:m.start()].strip()
            chapters.insert(i + 1, ChapterSummary(nxt, f"Chapter {nxt}", body, cs.char_offset + m.start()))
            print(f"[ThesisSifu] Split unheaded Chapter {nxt} out of Chapter {cs.chapter_num}")
        i += 1

_TOC_LINE = re.compile(r'(?:\t|\.{3,}|\s{2,})\s*\d{1,3}\s*$')

def _recover_chapter_one(text: str, chapters: list) -> None:
    """If Chapter 1's body heading wasn't recognised, only its Table-of-Contents line
    survives (a few hundred chars) and the real Chapter 1 is lost. Rebuild it from
    the first body "1.1 ..." heading up to the start of Chapter 2."""
    ch1 = next((c for c in chapters if c.chapter_num == "1"), None)
    ch2 = next((c for c in chapters if c.chapter_num == "2"), None)
    if not ch2 or (ch1 and len(ch1.text) >= 2000):
        return
    for m in re.finditer(r'(?:^|\n)[ \t]*1\.1[ \t]+\S[^\n]*', text[:ch2.char_offset]):
        line = m.group(0).strip()
        following = text[m.end():m.end() + 600].split("\n")[1:6]
        if _TOC_LINE.search(line) or sum(bool(_TOC_LINE.search(x)) for x in following) >= 2:
            continue                                   # still inside the ToC
        start = m.start()
        body = text[start:ch2.char_offset].strip()
        if len(body) < 2000:
            continue
        if ch1:
            ch1.text, ch1.char_offset = body, start
        else:
            chapters.insert(0, ChapterSummary("1", "Introduction", body, start))
        print(f"[ThesisSifu] Recovered Chapter 1 body ({len(body):,} chars)")
        return

def split_subsections(chapter: ChapterSummary) -> list[Subsection]:
    """Split a chapter into subsections.
    The _SUB_PAT lookahead (space+letter) already filters statistics,
    but we also enforce a minimum text length so tiny fragments don't
    consume Gemini quota returning empty results."""
    MIN_SUB_CHARS = 300   # skip subsections shorter than this
    matches = list(_SUB_PAT.finditer(chapter.text))
    if not matches:
        return [Subsection(
            chapter_num=chapter.chapter_num, subsection_num="—",
            title=chapter.chapter_title, text=chapter.text,
            char_offset=chapter.char_offset,
            expected_purpose=_purpose_for(chapter.chapter_title))]
    subs = []
    for i, m in enumerate(matches):
        sub_num   = m.group(1).strip()
        sub_title = (m.group(2) or "").strip() or f"Section {sub_num}"
        start     = m.start()
        end       = matches[i+1].start() if i+1 < len(matches) else len(chapter.text)
        sub_text  = chapter.text[start:end].strip()
        if len(sub_text) < MIN_SUB_CHARS:
            continue   # skip navigation entries, TOC lines, one-liner sections
        subs.append(Subsection(
            chapter_num=chapter.chapter_num, subsection_num=sub_num,
            title=sub_title, text=sub_text,
            char_offset=chapter.char_offset + start,
            expected_purpose=_purpose_for(sub_title),
        ))
    # If all were too short, fall back to treating whole chapter as one subsection
    if not subs:
        return [Subsection(
            chapter_num=chapter.chapter_num, subsection_num="—",
            title=chapter.chapter_title, text=chapter.text,
            char_offset=chapter.char_offset,
            expected_purpose=_purpose_for(chapter.chapter_title))]
    return subs


# ── Page estimator (uses full_text offset, not sub slice) ─────
def _estimate_page(full_text: str, excerpt: str, char_offset: int,
                   words_per_page: int = 350) -> int:
    """Find excerpt in full_text starting from char_offset; fall back to offset-based."""
    search_start = max(0, char_offset - 200)
    pos = full_text.find(excerpt[:50].strip(), search_start)
    if pos < 0:
        # Try shorter excerpt
        pos = full_text.find(excerpt[:25].strip(), search_start)
    if pos > 0:
        word_count = len(full_text[:pos].split())
        return max(1, word_count // words_per_page + 1)
    # Fall back to offset
    word_count = len(full_text[:char_offset].split())
    return max(1, word_count // words_per_page + 1)


def _find_content_start(text: str) -> int:
    """Skip over preamble pages (cover, admin forms, acknowledgements, table of contents).
    Returns the char position where substantive academic content begins.
    Looks for the Abstract, then Chapter 1, then any 'Introduction' heading."""
    markers = ["abstract\n", "Abstract\n", "ABSTRACT\n",
               "ABSTRAK\n", "Abstrak\n",
               "CHAPTER 1\n", "Chapter 1\n", "CHAPTER ONE\n",
               "BAB 1\n", "BAB SATU\n", "Bab 1\n",
               "1.1 Introduction\n", "1.1 Background\n", "1.1 Pengenalan\n",
               "Introduction\n1.", "INTRODUCTION\n", "PENGENALAN\n"]
    best = len(text)
    for marker in markers:
        pos = text.find(marker)
        if 0 < pos < best:
            best = pos
    # If nothing found in reasonable range, start at 0
    return best if best < len(text) * 0.3 else 0

_NUMBERED_HEADING = re.compile(r'^\s*\d{1,2}\.\d{1,3}(?:\.\d{1,3})?\s+\S')

def _lines_after_heading(text: str, headings: tuple, keep, limit: int = 8) -> list:
    """Items listed under the LAST body occurrence of a heading (e.g. 'Persoalan Kajian').
    Word list numbering is lost in extraction, so RQs/ROs arrive as bare lines; we
    collect lines that pass `keep` until the next numbered heading."""
    pat = re.compile(r'^[ \t]*(?:\d{1,2}\.\d{1,2}[ \t]+)?(?:' + "|".join(headings) + r')[ \t]*:?[ \t]*$',
                     re.IGNORECASE | re.MULTILINE)
    heads = [m for m in pat.finditer(text)]
    if not heads:
        return []
    # Re-join items that Word wrapped onto several lines (no closing punctuation yet).
    lines, buf = [], ""
    for line in text[heads[-1].end():].split("\n")[1:40]:
        s = line.strip()
        if not s:
            continue
        if _NUMBERED_HEADING.match(s) or re.match(r'^(BAB|CHAPTER)\s', s, re.IGNORECASE):
            if buf: lines.append(buf)
            buf = ""
            lines.append(None)                     # stop marker
            break
        buf = f"{buf} {s}".strip()
        if re.search(r'(?:[?.:;]|;\s*(?:dan|and|atau|or))\s*\.?$', buf, re.IGNORECASE):
            lines.append(buf); buf = ""
    if buf:
        lines.append(buf)
    out = []
    for s in lines:
        if s is None:
            break
        s = re.sub(r'^(?:\(?[a-z0-9ivx]{1,4}[\.\)])\s+', '', s, flags=re.IGNORECASE)
        if keep(s):
            out.append(s)
        if len(out) >= limit:
            break
    return out

def _extract_spine_fallback(text: str) -> dict:
    """Rule-based spine extraction as fallback when Gemini fails.
    Searches for explicit RQ/RO patterns, methodology keywords, etc."""
    d = {
        "title": "UNKNOWN", "discipline": "UNKNOWN",
        "problem_statement": "NOT FOUND", "research_gap": "NOT FOUND",
        "research_questions": [], "research_objectives": [], "hypotheses": [],
        "theory_used": "NOT FOUND", "variables": [],
        "methodology": "NOT FOUND", "sampling": "NOT FOUND",
        "instrument": "NOT FOUND", "analysis_technique": "NOT FOUND",
        "key_findings": [], "conclusions": []
    }
    # Title: look for ALL-CAPS line or explicit "Title:" field
    title_m = re.search(r'(?:Tajuk|Title)\s*[:：]\s*([^\n]{10,120})', text, re.IGNORECASE)
    if title_m: d["title"] = title_m.group(1).strip()
    else:
        # Try bold-looking ALL CAPS heading near top
        caps_m = re.search(r'\n([A-Z][A-Z\s]{15,80})\n', text[:3000])
        if caps_m: d["title"] = caps_m.group(1).strip()

    # Research Questions: look for RQ1/RQ2 or "Research Question 1"
    rq_matches = re.findall(
        r'RQ\s*\d+\s*[:：]\s*([^\n?]+\??)', text, re.IGNORECASE)
    if not rq_matches:
        rq_matches = re.findall(
            r'Research Question \d+\s*[:：]\s*([^\n?]+\??)', text, re.IGNORECASE)
    if not rq_matches:
        # BM: "Persoalan Kajian 1: ..." / "SK1: ..." / numbered "Adakah/Apakah/Sejauh manakah ...?"
        rq_matches = re.findall(
            r'(?:Persoalan Kajian|Soalan Kajian|SK)\s*\d+\s*[:：]\s*([^\n?]+\??)', text, re.IGNORECASE)
    if not rq_matches:
        rq_matches = _lines_after_heading(
            text, ("Persoalan Kajian", "Soalan Kajian", "Research Questions?"),
            keep=lambda s: s.rstrip(" .").endswith("?") and len(s) > 20)
    d["research_questions"] = [r.strip() for r in rq_matches[:6]]

    # Research Objectives: numbered "To verb..." / BM "Mengenal pasti / Mengkaji ..." lists
    ro_matches = re.findall(
        r'\d+\.\s+(To\s+(?:identify|examine|determine|explore|assess|investigate|describe|compare|evaluate|analyse|analyze)[^\n\.]{10,200})',
        text, re.IGNORECASE)
    if not ro_matches:
        ro_matches = _lines_after_heading(
            text, ("Objektif Kajian", "Research Objectives?"),
            keep=lambda s: bool(re.match(
                r'(?:To\s+\w+|(?:Untuk\s+)?(?:mengenal\s*pasti|mengenalpasti|mengkaji|menentukan|meneroka|'
                r'menilai|menganalisis|meninjau|membandingkan|menghuraikan|meneliti|mengukur))',
                s, re.IGNORECASE)) and len(s) > 20)
    d["research_objectives"] = [r.strip() for r in ro_matches[:6]]

    # Hypotheses (H1:, H01:, Ho1:)
    h_matches = re.findall(r'H[0o1]?\d+\s*[:：]\s*([^\n\.]{10,200})', text)
    d["hypotheses"] = [h.strip() for h in h_matches[:6]]

    # Methodology keywords
    if re.search(r'mixed[\s-]method', text, re.IGNORECASE):
        d["methodology"] = "Mixed-methods"
        if re.search(r'explanatory sequential', text, re.IGNORECASE):
            d["methodology"] = "Explanatory sequential mixed-methods"
        elif re.search(r'convergent parallel', text, re.IGNORECASE):
            d["methodology"] = "Convergent parallel mixed-methods"
    elif re.search(r'\bqualitative\b', text, re.IGNORECASE) and not re.search(r'\bquantitative\b', text, re.IGNORECASE):
        d["methodology"] = "Qualitative"
    elif re.search(r'\bquantitative\b', text, re.IGNORECASE):
        d["methodology"] = "Quantitative"

    # Analysis technique
    for technique in ["PLS-SEM", "CB-SEM", "thematic analysis", "Thematic Analysis",
                       "Pearson correlation", "regression", "ANOVA", "IPA",
                       "grounded theory", "Grounded Theory", "descriptive statistics"]:
        if technique.lower() in text.lower():
            existing = d["analysis_technique"]
            if existing == "NOT FOUND":
                d["analysis_technique"] = technique
            elif technique.lower() not in existing.lower():
                d["analysis_technique"] = existing + " + " + technique
            break

    # Theory
    theory_patterns = [
        ("Social Cognitive Theory", r"Social Cognitive Theory|Bandura.*self.efficacy"),
        ("TAM", r"\bTAM\b|Technology Acceptance Model"),
        ("UTAUT", r"\bUTAUT\b"),
        ("TPB", r"Theory of Planned Behaviour|Theory of Planned Behavior"),
        ("SDT", r"Self.Determination Theory"),
    ]
    for name, pat in theory_patterns:
        if re.search(pat, text, re.IGNORECASE):
            d["theory_used"] = name
            break

    # Sampling
    samp_m = re.search(r'(n\s*=\s*\d+|sample\s+of\s+\d+|\d+\s+(?:respondents|participants|teachers|students))',
                        text, re.IGNORECASE)
    if samp_m: d["sampling"] = samp_m.group(0).strip()

    # Abstract for problem + findings
    abs_m = re.search(r'Abstract\s*\n(.{200,2000}?)(?:\n\n|\nKeywords|\nAbstrak)', text, re.IGNORECASE | re.DOTALL)
    if abs_m:
        abstract = abs_m.group(1).strip()
        d["problem_statement"] = abstract[:400]
        # Findings often in last sentence of abstract
        sentences = [s.strip() for s in abstract.split('.') if len(s.strip()) > 30]
        if sentences:
            d["key_findings"] = [sentences[-1]] if sentences else []

    # Discipline from keywords
    for disc, pats in [
        ("Education", ["teacher", "student", "classroom", "learning", "instruction", "pedagogy"]),
        ("Information Systems", ["system", "technology", "adoption", "TAM", "IS"]),
        ("Business", ["firm", "organization", "management", "strategy"]),
        ("Health", ["patient", "health", "clinical", "medical"]),
    ]:
        if any(p.lower() in text[:5000].lower() for p in pats):
            d["discipline"] = disc
            break

    print(f"[ThesisSifu] Fallback spine: RQs={len(d['research_questions'])}, "
          f"ROs={len(d['research_objectives'])}, method={d['methodology']}, "
          f"theory={d['theory_used']}")
    return d


# ── Stage 0: Spine Extraction ──────────────────────────────────
_SPINE_PROMPT = """\
You are extracting the STRUCTURAL SPINE of an academic thesis or article.
Return STRICT JSON only — no markdown, no commentary, no preamble.

Fields (use "NOT FOUND" or [] if absent from the excerpt):
{
  "title": "exact title",
  "discipline": "best guess discipline (e.g., Information Systems, Marketing, Education)",
  "problem_statement": "1-3 sentence summary of the problem",
  "research_gap": "1-2 sentence summary of the gap claimed",
  "research_questions": ["RQ1 verbatim or paraphrased", "RQ2..."],
  "research_objectives": ["RO1...", "RO2..."],
  "hypotheses": ["H1...", "H2..."],
  "theory_used": "name of underpinning theory/framework if stated",
  "variables": ["IV1", "IV2", "DV", "MV (moderator/mediator)"],
  "methodology": "design + paradigm (e.g., quantitative cross-sectional survey)",
  "sampling": "technique + size if stated",
  "instrument": "questionnaire / interview protocol / etc.",
  "analysis_technique": "PLS-SEM / thematic / regression / etc.",
  "key_findings": ["finding 1", "finding 2"],
  "conclusions": ["conclusion 1", "conclusion 2"]
}

Document excerpt (front + back samples):
\"\"\"EXCERPT_PLACEHOLDER\"\"\"\
"""

_RQ_ANCHORS = ("RQ1", "Research Question 1", "Research Questions", "RESEARCH QUESTIONS",
               "Persoalan Kajian", "PERSOALAN KAJIAN", "Objektif Kajian", "OBJEKTIF KAJIAN",
               "Research Objectives", "RESEARCH OBJECTIVES")

def _spine_sample(text: str, chs: list) -> str:
    """Build the excerpt Gemini reads to extract the spine.
    With real chapters: all of Chapter 1's front, the opening of each middle
    chapter (method / results), and the opening of the LAST chapter, where the
    conclusions are. (text[-3000:] is usually the reference list or appendices.)"""
    real = [c for c in chs if not c.chapter_title.startswith("Section ")]
    if len(real) >= 3:
        parts = [f"[CHAPTER {real[0].chapter_num}]\n" + real[0].text[:18000]]
        for c in real[1:-1][:4]:
            parts.append(f"[CHAPTER {c.chapter_num} — opening]\n" + c.text[:5000])
        parts.append(f"[CHAPTER {real[-1].chapter_num} — final chapter]\n" + real[-1].text[:12000])
        return "\n\n[...]\n\n".join(parts)[:45000]

    content_start = _find_content_start(text)
    content = text[content_start:]
    front = content[:8000]
    rq_pos = min((p for p in (text.find(a) for a in _RQ_ANCHORS) if p > 0), default=-1)
    middle = text[max(0, rq_pos - 200):rq_pos + 3000] if rq_pos > 0 else content[8000:12000]
    concl = max(text.rfind("CONCLUSION"), text.rfind("KESIMPULAN"), text.rfind("Conclusion"))
    back = text[concl:concl + 6000] if concl > len(text) * 0.5 else text[-6000:]
    return (front + "\n\n[...RQ section...]\n\n" + middle +
            "\n\n[...conclusion section...]\n\n" + back)[:30000]

async def extract_spine(text: str, chs: list) -> ThesisSpine:
    prompt = _SPINE_PROMPT.replace("EXCERPT_PLACEHOLDER", _spine_sample(text, chs))

    d = None
    if text.strip():
        d = await _gemini_json(prompt, "spine_extraction")
        if not isinstance(d, dict):
            d = None
        else:
            print(f"[ThesisSifu] Gemini spine: RQs={len(d.get('research_questions',[]))}")

    # ── Fallback: rule-based extraction if Gemini failed or returned mostly NOT FOUND ──
    if d is None:
        print("[ThesisSifu] Spine: using rule-based fallback (no Gemini result)")
        d = _extract_spine_fallback(text)
    else:
        # Merge: fill in any NOT FOUND fields with rule-based results
        fallback = _extract_spine_fallback(text)
        for key in ["title", "problem_statement", "research_questions", "research_objectives",
                    "hypotheses", "theory_used", "methodology", "analysis_technique",
                    "sampling", "key_findings"]:
            gemini_val = d.get(key)
            fb_val = fallback.get(key)
            if (not gemini_val or gemini_val in ("UNKNOWN","NOT FOUND",[])) and fb_val and fb_val not in ("UNKNOWN","NOT FOUND",[]):
                d[key] = fb_val
                print(f"[ThesisSifu] Spine fallback filled: {key}")

    # ── Coerce every field to its declared type before constructing the dataclass ──
    # Gemini's JSON isn't schema-enforced: a field documented as a string can come
    # back as a list/dict/number, and a field documented as a list can come back
    # as a bare string. Either mismatch used to propagate silently (or crash later
    # in .replace()/.join() calls deep in run_alignment_audit/run_examiner). Both
    # helpers are idempotent, so this is safe regardless of whether `d` came from
    # Gemini or the rule-based fallback above.
    return ThesisSpine(
        title=              _flatten_to_str(d.get("title","UNKNOWN")),
        discipline=         _flatten_to_str(d.get("discipline","UNKNOWN")),
        problem_statement=  _flatten_to_str(d.get("problem_statement","NOT FOUND")),
        research_gap=       _flatten_to_str(d.get("research_gap","NOT FOUND")),
        research_questions= _flatten_to_list(d.get("research_questions",[])),
        research_objectives=_flatten_to_list(d.get("research_objectives",[])),
        hypotheses=         _flatten_to_list(d.get("hypotheses",[])),
        theory_used=        _flatten_to_str(d.get("theory_used","NOT FOUND")),
        variables=          _flatten_to_list(d.get("variables",[])),
        methodology=        _flatten_to_str(d.get("methodology","NOT FOUND")),
        sampling=           _flatten_to_str(d.get("sampling","NOT FOUND")),
        instrument=         _flatten_to_str(d.get("instrument","NOT FOUND")),
        analysis_technique= _flatten_to_str(d.get("analysis_technique","NOT FOUND")),
        key_findings=       _flatten_to_list(d.get("key_findings",[])),
        conclusions=        _flatten_to_list(d.get("conclusions",[])),
    )


# ── Stage 1: Classification ─────────────────────────────────────
_CLS_PROMPT = """\
Classify this academic document. Choose ONE:
JOURNAL_ARTICLE | UNDERGRADUATE | MASTERS | PHD

Return ONLY valid JSON (no markdown):
{"type":"PHD","confidence":"HIGH","signals":["s1","s2"],"title":"t","authors":"a","field":"f","institution":"i"}

Excerpt:
\"\"\"EXCERPT_PLACEHOLDER\"\"\"\
"""

async def classify_document(text: str) -> dict:
    default = {"type":"MASTERS","confidence":"LOW","signals":["fallback"],
               "title":"UNKNOWN","authors":"UNKNOWN","field":"UNKNOWN","institution":"UNKNOWN"}
    if not gemini_client or not text.strip():
        return default
    prompt = _CLS_PROMPT.replace("EXCERPT_PLACEHOLDER", text[:5000])
    d = await _gemini_json(prompt, "classify_document")
    if not isinstance(d, dict):
        return default
    return {**default, **d}


# ── Stage 2b: Chapter digests (the golden-thread map) ─────────
_DIGEST_PROMPT = """\
You are reading ONE chapter of a thesis to record what it COMMITS to, so an examiner can
later check the golden thread across chapters (Problem -> Gap -> RQ/RO/H -> Theory/Constructs
-> Design/Instrument -> Analysis -> Findings -> Discussion -> Conclusion/Contribution).

CHAPTER CHNUM_PLACEHOLDER — CHTITLE_PLACEHOLDER

Return STRICT JSON only:
{
  "chapter_role": "introduction|literature_review|methodology|results|discussion|conclusion|other",
  "summary": "3-5 sentences in English on what this chapter actually does",
  "commitments": [
    {"type": "problem|gap|research_question|objective|hypothesis|theory|construct|conceptual_framework|design|sample|instrument|measurement|analysis|finding|interpretation|contribution|implication|limitation|recommendation|conclusion",
     "label": "short id, e.g. RQ2, H3, Self-efficacy, PLS-SEM",
     "statement": "what the chapter says, one sentence in English",
     "quote": "verbatim supporting text from the chapter, max 200 characters, original language",
     "section": "section number if visible (e.g. 3.4.2), else empty"}
  ],
  "constructs": ["every variable / construct / theme named in this chapter, worded exactly as in the text"]
}

Record at most 30 commitments. Prioritise: RQs, objectives, hypotheses, the underpinning theory,
each construct and HOW it is measured (instrument, items, source), the analysis used for each
RQ/hypothesis, the finding for each RQ/hypothesis, and claimed contributions. Record only what
the text says; do not judge quality here.

Chapter text:
\"\"\"TEXT_PLACEHOLDER\"\"\"\
"""

async def digest_chapter(cs: ChapterSummary) -> None:
    prompt = (_DIGEST_PROMPT
              .replace("CHNUM_PLACEHOLDER", cs.chapter_num)
              .replace("CHTITLE_PLACEHOLDER", cs.chapter_title[:80])
              .replace("TEXT_PLACEHOLDER", cs.text[:150000]))
    d = await _gemini_json(prompt, f"digest_ch{cs.chapter_num}")
    if isinstance(d, dict):
        d["commitments"] = [c for c in _flatten_to_list(d.get("commitments")) if isinstance(c, dict)][:30]
        d["constructs"] = [str(x) for x in _flatten_to_list(d.get("constructs"))][:40]
        cs.digest = d
        print(f"[ThesisSifu] Digest ch{cs.chapter_num}: role={d.get('chapter_role')} "
              f"commitments={len(d['commitments'])}")

async def digest_chapters(chs: list) -> None:
    sem = asyncio.Semaphore(6)
    async def one(cs):
        async with sem:
            await digest_chapter(cs)
    await asyncio.gather(*(one(cs) for cs in chs if cs.text.strip()))

_THREAD_TYPES = ("research_question", "objective", "hypothesis", "theory", "construct",
                 "conceptual_framework", "instrument", "measurement", "analysis",
                 "finding", "contribution", "conclusion")

def thread_context(chs: list, limit: int = 4000) -> str:
    """Compact cross-chapter commitments, given to every subsection auditor so it
    can judge a paragraph against what OTHER chapters promise or report."""
    lines = []
    for cs in chs:
        for c in cs.digest.get("commitments", []):
            if c.get("type") in _THREAD_TYPES:
                lines.append(f"Ch{cs.chapter_num} {c.get('type')} {c.get('label','')}: "
                             f"{_shorten(str(c.get('statement','')), 160)}")
    return "\n".join(lines)[:limit] or "(not available)"

def digests_block(chs: list, limit: int = 70000) -> str:
    """Full digest text for the alignment audit."""
    out = []
    for cs in chs:
        d = cs.digest
        if not d:
            out.append(f"=== Chapter {cs.chapter_num} — {cs.chapter_title} === (digest unavailable)")
            continue
        out.append(f"=== Chapter {cs.chapter_num} — {cs.chapter_title} (role: {chapter_role(cs)}) ===")
        out.append(f"Summary: {d.get('summary','')}")
        for c in d.get("commitments", []):
            sec = f" (§{c['section']})" if c.get("section") else ""
            out.append(f"- [{c.get('type','')}] {c.get('label','')}: {c.get('statement','')}{sec}"
                       f"\n    quote: \"{_shorten(str(c.get('quote','')), 200)}\"")
        if d.get("constructs"):
            out.append("Constructs named: " + "; ".join(d["constructs"]))
    return "\n".join(out)[:limit]


# ── Stage 2: Subsection paragraph audit ───────────────────────
# CANONICAL libraries removed from this prompt — they added ~4000 chars to every
# one of 70+ parallel calls, blowing the total processing time past Railway's timeout.
# Gemini still recommends frameworks/methods by name from its training knowledge.
_SUBSECTION_PROMPT_TEMPLATE = """\
You are a rigorous but fair thesis supervisor auditing one SUBSECTION of a DOC_TYPE_PLACEHOLDER.
Your feedback must be specific enough that the student knows exactly what to change.

LANG_PLACEHOLDER

THESIS SPINE:
  Title:      TITLE_PLACEHOLDER
  Problem:    PROBLEM_PLACEHOLDER
  Gap:        GAP_PLACEHOLDER
  RQs:        RQS_PLACEHOLDER
  ROs:        ROS_PLACEHOLDER
  Hypotheses: HYPS_PLACEHOLDER
  Theory:     THEORY_PLACEHOLDER
  Method:     METHOD_PLACEHOLDER
  Analysis:   ANALYSIS_PLACEHOLDER

WHAT THE OTHER CHAPTERS COMMIT TO (golden-thread map — use it to catch cross-chapter breaks):
THREAD_PLACEHOLDER

CHAPTER CHNUM_PLACEHOLDER — CHTITLE_PLACEHOLDER (role: ROLE_PLACEHOLDER)
SUBSECTION SUBNUM_PLACEHOLDER — SUBTITLE_PLACEHOLDER
Expected purpose: PURPOSE_PLACEHOLDER
What examiners check in this kind of chapter:
CHECKS_PLACEHOLDER
CURRENT YEAR: YEAR_PLACEHOLDER (Do not flag 2024-YEAR_PLACEHOLDER citations as "future" or "unpublished").

Report 0-5 GENUINE issues — only real weaknesses an experienced examiner would raise. Do not pad:
if the subsection is sound, return []. Prefer substantive issues (argument, evidence, alignment,
method, interpretation) over wording, and do not repeat an issue that belongs to another section.

Write as a senior supervisor speaking to the student about THEIR text: refer to their own
headings, research questions, constructs and page content. Never mention this tool's internal
terms ("thesis spine", "expected purpose", "golden-thread map", "subsection audit", "placeholder").
The spine and map above were extracted automatically and may be incomplete: a field reading
NOT FOUND means the extraction missed it, NOT that the thesis lacks it. Only say something is
missing from the thesis if the text you are reading shows that it is.
Return a JSON array; each element:
{
  "para_excerpt": "first 80 characters of the paragraph, copied verbatim",
  "severity": "CRITICAL|MODERATE|SUGGESTION",
  "issue": "what is wrong AND why it matters to an examiner; name the exact claim, construct or RQ involved",
  "golden_thread_link": "the cross-chapter link affected, e.g. 'RQ2 -> Ch3 instrument' or 'Ch1 gap -> Ch2 synthesis'; empty if none",
  "recommendation": "the concrete fix: what to add, remove or restructure, with a short example sentence or structure where useful",
  "literature_needed": "the kind of evidence needed (e.g. 'recent empirical studies on X among Y in Malaysia')",
  "search_query": "4-8 English keywords to find peer-reviewed literature supporting the fix; empty if no literature is needed",
  "theory_needed": "named theory/framework or none",
  "suggested_framework": "e.g. TAM (Davis, 1989) or empty string",
  "suggested_method": "e.g. PLS-SEM (Hair et al., 2022) or empty string"
}

Severity (calibrate like an external examiner — most issues are MODERATE; CRITICAL is rare and
means the examiner would require it fixed before the thesis can pass):
  CRITICAL   — breaks the golden thread or validity: contradicts/ignores an RQ, RO or hypothesis;
               a construct is not measured or not analysed; the method cannot answer the RQ; a
               finding maps to no RQ; a conclusion overreaches the data; the subsection fails its purpose.
  MODERATE   — weakens the argument: unsupported claim, descriptive rather than critical synthesis,
               outdated or thin evidence, missing justification or threshold, unclear operationalisation.
  SUGGESTION — would strengthen an already acceptable section.

Never invent references, titles or DOIs; describe the evidence needed instead.
Return ONLY the JSON array.

Subsection text:
\"\"\"TEXT_PLACEHOLDER\"\"\"\
"""

async def audit_subsection(sub: Subsection, cs: ChapterSummary, doc_type: str,
                            spine: ThesisSpine, full_text: str, thread: str,
                            lang: str) -> list[ParagraphComment]:
    if not gemini_client or not sub.text.strip():
        return []
    role = chapter_role(cs)
    year = str(datetime.now().year)
    prompt = (
        _SUBSECTION_PROMPT_TEMPLATE
        .replace("DOC_TYPE_PLACEHOLDER", doc_type)
        .replace("LANG_PLACEHOLDER",     _lang_instruction(lang))
        .replace("TITLE_PLACEHOLDER",    spine.title[:150])
        .replace("PROBLEM_PLACEHOLDER",  spine.problem_statement[:300])
        .replace("GAP_PLACEHOLDER",      spine.research_gap[:250])
        .replace("RQS_PLACEHOLDER",      "; ".join(spine.research_questions[:6]) or "NOT FOUND")
        .replace("ROS_PLACEHOLDER",      "; ".join(spine.research_objectives[:6]) or "NOT FOUND")
        .replace("HYPS_PLACEHOLDER",     "; ".join(spine.hypotheses[:8]) or "none stated")
        .replace("THEORY_PLACEHOLDER",   spine.theory_used[:150])
        .replace("METHOD_PLACEHOLDER",   spine.methodology[:150])
        .replace("ANALYSIS_PLACEHOLDER", spine.analysis_technique[:150])
        .replace("THREAD_PLACEHOLDER",   thread)
        .replace("CHNUM_PLACEHOLDER",    sub.chapter_num)
        .replace("CHTITLE_PLACEHOLDER",  cs.chapter_title[:60])
        .replace("ROLE_PLACEHOLDER",     role)
        .replace("SUBNUM_PLACEHOLDER",   sub.subsection_num)
        .replace("SUBTITLE_PLACEHOLDER", sub.title[:60])
        .replace("PURPOSE_PLACEHOLDER",  sub.expected_purpose[:150])
        .replace("CHECKS_PLACEHOLDER",   _CHAPTER_CHECKS.get(role, _CHAPTER_CHECKS["other"]))
        .replace("YEAR_PLACEHOLDER",     year)
        .replace("TEXT_PLACEHOLDER",     sub.text[:60000])
    )
    ctx = f"ch{sub.chapter_num}.{sub.subsection_num}"
    items = await _gemini_json(prompt, f"audit_subsection_{ctx}")
    if items is None:
        return []
    if isinstance(items, dict):     # a single object instead of an array
        items = [items]
    comments = []
    for idx, it in enumerate(x for x in items if isinstance(x, dict)):
        excerpt = str(it.get("para_excerpt",""))[:150]
        sev = str(it.get("severity","MODERATE")).upper()
        comments.append(ParagraphComment(
            chapter=            f"Chapter {sub.chapter_num}",
            subsection=         sub.subsection_num,
            subsection_title=   sub.title,
            para_index=         idx,
            page_estimate=      _estimate_page(full_text, excerpt, sub.char_offset),
            para_excerpt=       excerpt,
            severity=           sev if sev in SEV_COLORS else "MODERATE",
            issue=              _flatten_to_str(it.get("issue","")),
            recommendation=     _flatten_to_str(it.get("recommendation","")),
            literature_needed=  _flatten_to_str(it.get("literature_needed","")),
            theory_needed=      _flatten_to_str(it.get("theory_needed","none")),
            suggested_framework=_flatten_to_str(it.get("suggested_framework","")),
            suggested_method=   _flatten_to_str(it.get("suggested_method","")),
            golden_thread_link= _flatten_to_str(it.get("golden_thread_link","")),
            search_query=       _flatten_to_str(it.get("search_query","")),
        ))
    print(f"[ThesisSifu] {ctx}: {len(comments)} comments")
    return comments


# ── Stage 3b: Literature grounding (real, retrievable works only) ──
# The model never writes a reference. We search a scholarly index with the
# comment's keywords, then the model may only PICK from the records returned,
# so every reference in the reports exists, with a DOI or URL.
def _http_get_json(url: str, timeout: float = 15.0, retries: int = 2):
    """GET JSON, backing off on HTTP 429 (Crossref's public pool rate-limits hard)."""
    import time, urllib.error
    headers = {"User-Agent": "ThesisSifuPro/4.2" + (f" (mailto:{LITERATURE_MAILTO})" if LITERATURE_MAILTO else "")}
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code != 429 or attempt == retries:
                raise
            wait = e.headers.get("Retry-After", "")
            time.sleep(min(float(wait), 10.0) if wait.replace(".", "", 1).isdigit() else 2.0 * (attempt + 1))

def _openalex_abstract(inv) -> str:
    if not isinstance(inv, dict):
        return ""
    pos = sorted((p, w) for w, ps in inv.items() for p in ps)
    return " ".join(w for _, w in pos)

def _search_openalex(query: str, n: int) -> list:
    params = {
        "search": query, "per_page": n,
        "filter": "from_publication_date:2010-01-01,type:article|review|book-chapter",
        "select": "id,doi,display_name,publication_year,cited_by_count,authorships,"
                  "primary_location,abstract_inverted_index",
        "api_key": OPENALEX_API_KEY,
    }
    d = _http_get_json("https://api.openalex.org/works?" + urllib.parse.urlencode(params))
    out = []
    for w in d.get("results", []):
        src = ((w.get("primary_location") or {}).get("source") or {})
        out.append({
            "title":   w.get("display_name") or "",
            "year":    w.get("publication_year"),
            "authors": [_split_name(a.get("author", {}).get("display_name", "")) for a in w.get("authorships", [])],
            "venue":   src.get("display_name") or "",
            "doi":     (w.get("doi") or "").replace("https://doi.org/", ""),
            "url":     w.get("doi") or w.get("id") or "",
            "cited_by": w.get("cited_by_count") or 0,
            "abstract": _openalex_abstract(w.get("abstract_inverted_index")),
        })
    return out

def _search_crossref(query: str, n: int) -> list:
    params = {
        "query.bibliographic": query, "rows": n,
        "filter": "from-pub-date:2010,type:journal-article",
        "select": "DOI,title,author,issued,container-title,is-referenced-by-count,abstract",
    }
    if LITERATURE_MAILTO:
        params["mailto"] = LITERATURE_MAILTO
    d = _http_get_json("https://api.crossref.org/works?" + urllib.parse.urlencode(params))
    out = []
    for w in d.get("message", {}).get("items", []):
        parts = (w.get("issued") or {}).get("date-parts") or [[None]]
        out.append({
            "title":   (w.get("title") or [""])[0],
            "year":    parts[0][0] if parts and parts[0] else None,
            "authors": [(a.get("family", ""), a.get("given", "")) for a in w.get("author", []) if a.get("family")],
            "venue":   (w.get("container-title") or [""])[0],
            "doi":     w.get("DOI") or "",
            "url":     f"https://doi.org/{w['DOI']}" if w.get("DOI") else "",
            "cited_by": w.get("is-referenced-by-count") or 0,
            "abstract": re.sub(r"<[^>]+>", " ", w.get("abstract") or ""),
        })
    return out

async def search_literature(query: str, n: int = 6) -> list:
    """Real works for a query: OpenAlex when keyed, Crossref otherwise (each falls back to the other)."""
    order = [_search_openalex, _search_crossref] if OPENALEX_API_KEY else [_search_crossref]
    for fn in order:
        try:
            works = await asyncio.to_thread(fn, query, n)
            works = [w for w in works if w["title"] and w["year"] and w["url"]]
            if works:
                return works
        except Exception as e:
            print(f"[ThesisSifu] Literature search via {fn.__name__} failed for '{query}': {e}")
    return []

_NAME_PARTICLES = {"de", "da", "di", "van", "von", "der", "den", "del", "la", "le", "du", "dos"}

def _split_name(display: str) -> tuple:
    """'Debbie De Neve' -> ('De Neve', 'Debbie'). Used when only a display name exists."""
    parts = display.split()
    if not parts:
        return ("", "")
    i = len(parts) - 1
    while i > 0 and parts[i - 1].lower() in _NAME_PARTICLES:
        i -= 1
    return (" ".join(parts[i:]), " ".join(parts[:i]))

def format_citation(w: dict) -> str:
    """APA-style citation from a retrieved record."""
    def short(name: tuple) -> str:
        family, given = name
        initials = " ".join(p[0] + "." for p in re.split(r"[\s\-]+", given) if p)
        return f"{family}, {initials}".strip().rstrip(",")
    names = [short(a) for a in w.get("authors", []) if a and a[0]][:20]
    if not names:
        auth = "Anon."
    elif len(names) == 1:
        auth = names[0]
    elif len(names) <= 3:
        auth = ", ".join(names[:-1]) + ", & " + names[-1]
    else:
        auth = names[0] + " et al."
    venue = f" {w['venue']}." if w.get("venue") else ""
    link = f" https://doi.org/{w['doi']}" if w.get("doi") else (f" {w['url']}" if w.get("url") else "")
    return f"{auth} ({w.get('year')}). {w.get('title','').rstrip('.')}.{venue}{link}"

_LIT_SELECT_PROMPT = """\
You are helping a thesis supervisor attach REAL supporting literature to feedback comments.
For each comment you are given candidate works retrieved from a scholarly database.
Pick up to 3 candidates per comment that DIRECTLY help the student fix the issue (evidence,
definitions, measurement, method, or theory). Prefer well-cited and recent (last ~10 years)
peer-reviewed work. If no candidate is genuinely relevant, pick none — never force a match.

LANG_PLACEHOLDER

Return STRICT JSON only:
{"selections": [{"comment_id": "c3", "picks": [{"candidate_id": "c3-2", "how_to_use": "one sentence: what to take from this source and where in the thesis to use it"}]}]}
Use only comment_id and candidate_id values shown below.

COMMENTS AND CANDIDATES
=======================
BLOCK_PLACEHOLDER\
"""

async def attach_literature(chs: list, lang: str) -> list:
    """Attach verified references to the most serious comments.
    Returns the de-duplicated list of works used (for the examiner report)."""
    targets = [c for cs in chs for c in cs.comments
               if c.severity in ("CRITICAL", "MODERATE") and c.search_query.strip()]
    targets.sort(key=lambda c: c.severity != "CRITICAL")
    targets = targets[:MAX_LIT_COMMENTS]
    if not targets:
        return []

    # One search per distinct query. Crossref's public pool allows little
    # concurrency, so it is searched one query at a time; OpenAlex in parallel.
    sem = asyncio.Semaphore(4 if OPENALEX_API_KEY else 1)
    queries = list(dict.fromkeys(c.search_query.strip().lower() for c in targets))
    async def fetch(q):
        async with sem:
            return q, await search_literature(q)
    found = dict(await asyncio.gather(*(fetch(q) for q in queries)))
    results = [found.get(c.search_query.strip().lower(), []) for c in targets]

    blocks, candidates = [], {}
    for i, (c, works) in enumerate(zip(targets, results)):
        if not works:
            continue
        cid = f"c{i}"
        lines = [f"[{cid}] ISSUE: {_shorten(c.issue, 400)}",
                 f"      NEEDED: {_shorten(c.literature_needed, 200)}"]
        for j, w in enumerate(works):
            kid = f"{cid}-{j}"
            candidates[kid] = (c, w)
            lines.append(f"   ({kid}) {w['title']} ({w['year']}; {w['venue'] or 'n/a'}; cited {w['cited_by']}x)"
                         f" — {_shorten(w['abstract'], 350) or 'no abstract'}")
        blocks.append("\n".join(lines))
    if not blocks:
        print("[ThesisSifu] Literature: no search results")
        return []

    prompt = (_LIT_SELECT_PROMPT
              .replace("LANG_PLACEHOLDER", _lang_instruction(lang))
              .replace("BLOCK_PLACEHOLDER", "\n\n".join(blocks)))
    d = await _gemini_json(prompt, "literature_selection")
    used, seen = [], set()
    for sel in (d or {}).get("selections", []) if isinstance(d, dict) else []:
        for p in _flatten_to_list(sel.get("picks"))[:3]:
            if not isinstance(p, dict) or p.get("candidate_id") not in candidates:
                continue                      # ignore anything not actually retrieved
            c, w = candidates[p["candidate_id"]]
            if not p["candidate_id"].startswith(str(sel.get("comment_id", "")) + "-"):
                continue                      # a pick must belong to its own comment
            ref = {"citation": format_citation(w), "how_to_use": _flatten_to_str(p.get("how_to_use", ""))}
            c.references.append(ref)
            key = (w.get("doi") or w["title"]).lower()
            if key not in seen:
                seen.add(key)
                used.append(ref["citation"])
    print(f"[ThesisSifu] Literature: {sum(len(c.references) for c in targets)} references "
          f"attached to {sum(1 for c in targets if c.references)}/{len(targets)} comments")
    return used


# ── Stage 3: Alignment Audit (Claude) ─────────────────────────
_ALIGNMENT_SYSTEM = """\
You are a senior academic examiner auditing the STRUCTURAL ALIGNMENT (golden thread) of a thesis:
Problem -> Gap -> RQ -> RO -> Hypothesis -> Theory/Constructs -> Design/Instrument -> Analysis
-> Findings -> Discussion -> Conclusion/Contribution.
The full thesis is provided above; the chapter digests below are an index of what each chapter
commits to. Verify against the thesis text itself. A link is only ALIGNED if the text shows it in
BOTH chapters; if a later chapter never picks up what an earlier one promised (or the reverse),
that is a break. The spine was extracted automatically and may be incomplete — trust the thesis.
Cite evidence as chapter + section + short quote. Be specific and fair: say what is aligned as
well as what is broken.\
"""

_ALIGNMENT_SCHEMA = {
    "type": "object",
    "properties": {
        "rows": {"type": "array", "items": {
            "type": "object",
            "properties": {k: {"type": "string"} for k in (
                "rq", "ro", "hypothesis", "theory_construct", "method", "instrument",
                "analysis", "finding", "discussion", "conclusion", "broken_link",
                "evidence", "note", "fix")} | {
                "status": {"type": "string", "enum": ["ALIGNED", "PARTIAL", "MISALIGNED", "MISSING", "UNCLEAR"]}},
            "required": ["rq", "ro", "hypothesis", "theory_construct", "method", "instrument",
                         "analysis", "finding", "discussion", "conclusion", "status",
                         "broken_link", "evidence", "note", "fix"],
            "additionalProperties": False}},
        "chapter_links": {"type": "array", "items": {
            "type": "object",
            "properties": {k: {"type": "string"} for k in (
                "from_chapter", "to_chapter", "link", "evidence", "fix")} | {
                "status": {"type": "string", "enum": ["ALIGNED", "PARTIAL", "MISALIGNED", "MISSING", "UNCLEAR"]}},
            "required": ["from_chapter", "to_chapter", "link", "status", "evidence", "fix"],
            "additionalProperties": False}},
        "construct_consistency": {"type": "array", "items": {
            "type": "object",
            "properties": {k: {"type": "string"} for k in ("construct", "trace", "note")} | {
                "status": {"type": "string", "enum": ["ALIGNED", "PARTIAL", "MISALIGNED", "MISSING", "UNCLEAR"]}},
            "required": ["construct", "trace", "status", "note"],
            "additionalProperties": False}},
        "golden_thread_score": {"type": "string", "enum": ["STRONG", "ACCEPTABLE", "WEAK", "BROKEN"]},
        "overall_verdict": {"type": "string"},
        "critical_gaps": {"type": "array", "items": {"type": "string"}},
        "structural_recommendations": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["rows", "chapter_links", "construct_consistency", "golden_thread_score",
                 "overall_verdict", "critical_gaps", "structural_recommendations"],
    "additionalProperties": False,
}

_PANEL_SYSTEM = """\
You are a senior academic examiner with 25 years of experience supervising and examining
theses. You read the candidate's thesis yourself and judge it on its substance. Your feedback
is specific (chapter, section, quote), fair (credit genuine strengths), and actionable (say
exactly what to change). You are an intellectual critic, not a grammar checker.\
"""

def thesis_document(chs: list, limit: int = 400000) -> str:
    """The thesis chapters as one document for Claude (front matter excluded).
    Each chapter is capped so appendices pasted into the last chapter can't crowd out the rest."""
    parts = [f"<chapter number=\"{cs.chapter_num}\" title=\"{cs.chapter_title}\">\n{cs.text[:120000]}\n</chapter>"
             for cs in chs if cs.text.strip()]
    return "<thesis>\n" + "\n".join(parts)[:limit] + "\n</thesis>"

async def _claude_text(system: str, prompt: str, context: str,
                       schema: Optional[dict] = None, max_tokens: int = 32000,
                       document: str = "") -> str:
    """One Claude call (streamed, so long outputs don't hit HTTP timeouts).
    The thesis document goes first with a cache breakpoint: the alignment audit and the
    examiner share the same system prompt and document, so the second call reads it from cache.
    `system` holds the stage-specific instructions and is sent after the document."""
    content = []
    if document:
        content.append({"type": "text", "text": document, "cache_control": {"type": "ephemeral"}})
    content.append({"type": "text", "text": f"{system}\n\n{prompt}"})
    kwargs = dict(model=CLAUDE_MODEL, max_tokens=max_tokens, system=_PANEL_SYSTEM,
                  thinking={"type": "adaptive"},
                  messages=[{"role": "user", "content": content}])
    output_config = {"effort": CLAUDE_EFFORT}
    if schema:
        output_config["format"] = {"type": "json_schema", "schema": schema}
    kwargs["output_config"] = output_config
    async with claude_client.messages.stream(**kwargs) as stream:
        msg = await stream.get_final_message()
    if msg.stop_reason == "refusal":
        raise RuntimeError(f"{context}: model declined the request")
    if msg.stop_reason == "max_tokens":
        print(f"[ThesisSifu] WARNING {context}: hit max_tokens, output may be truncated")
    u = msg.usage
    print(f"[ThesisSifu] {context}: in={u.input_tokens} cache_read={getattr(u, 'cache_read_input_tokens', 0)} "
          f"cache_write={getattr(u, 'cache_creation_input_tokens', 0)} out={u.output_tokens} stop={msg.stop_reason}")
    return "".join(b.text for b in msg.content if b.type == "text").strip()

_ALIGNMENT_PROMPT_TEMPLATE = """\
LANG_PLACEHOLDER

THESIS SPINE (summary extracted from the thesis)
================================================
Title:                TITLE_PLACEHOLDER
Discipline:           DISCIPLINE_PLACEHOLDER
Problem statement:    PROBLEM_PLACEHOLDER
Research gap:         GAP_PLACEHOLDER
Research questions:   RQS_PLACEHOLDER
Research objectives:  ROS_PLACEHOLDER
Hypotheses:           HYPS_PLACEHOLDER
Theory used:          THEORY_PLACEHOLDER
Variables/constructs: VARS_PLACEHOLDER
Methodology:          METHOD_PLACEHOLDER
Sampling:             SAMPLING_PLACEHOLDER
Instrument:           INSTRUMENT_PLACEHOLDER
Analysis technique:   ANALYSIS_PLACEHOLDER
Key findings:         FINDINGS_PLACEHOLDER
Conclusions:          CONCLUSIONS_PLACEHOLDER

CHAPTER DIGESTS (what each chapter actually commits to, with verbatim quotes)
=============================================================================
DIGESTS_PLACEHOLDER

ISSUES RAISED BY THE SUBSECTION AUDIT
=====================================
CRITICAL_ISSUES_PLACEHOLDER

YOUR TASK
=========
1. rows — one row per research question (if there are no RQs, one per objective or hypothesis).
   Trace it through every chapter: matching RO and hypothesis; the theory/construct behind it;
   the method; the instrument items that measure its constructs; the analysis that answers it;
   the finding reported for it; how the discussion interprets it; the conclusion drawn.
   Use "NOT FOUND" for any link the digests do not show. Set status, name the first
   broken_link (e.g. "RQ3 -> Ch4 findings: no result reported"), give evidence
   (chapter, section, short quote), a note, and a concrete fix.
2. chapter_links — check these links between chapters, each with status, evidence and fix:
   Ch1 problem -> Ch1 gap; Ch1 gap -> Ch2 synthesis (does the review actually establish the gap?);
   Ch2 theory -> hypotheses / conceptual framework; Ch2 constructs -> Ch3 instrument;
   Ch3 analysis plan -> Ch4 analyses actually run; Ch4 findings -> Ch5 discussion (interpreted
   against Ch2 literature and theory); Ch5 conclusions/contributions -> Ch1 objectives and significance.
   Adapt chapter numbers to this thesis's structure.
3. construct_consistency — for every key construct/variable: trace where it is defined, measured,
   analysed and discussed (e.g. "Ch1 §1.6 defined; Ch2 §2.3 reviewed; Ch3 §3.5 measured (8 items);
   Ch4 Table 4.7 tested; Ch5 not discussed"). Flag renamed, dropped or added constructs.
4. golden_thread_score, overall_verdict (3-5 sentences), critical_gaps (most serious first,
   each specific), structural_recommendations (each a concrete action; name frameworks or
   methods with seminal authors where relevant).\
"""

async def run_alignment_audit(spine: ThesisSpine, chs: list[ChapterSummary], lang: str) -> AlignmentMatrix:
    if not claude_client:
        return AlignmentMatrix(
            overall_verdict="Alignment audit unavailable — ANTHROPIC_API_KEY not set.",
            golden_thread_score="UNCLEAR")
    issues = [f"[{c.severity}] {c.chapter} §{c.subsection}"
              f"{' (' + c.golden_thread_link + ')' if c.golden_thread_link else ''}: {_shorten(c.issue, 300)}"
              for cs in chs for c in cs.comments if c.severity in ("CRITICAL", "MODERATE")]
    issues.sort(key=lambda s: not s.startswith("[CRITICAL]"))
    issue_block = "\n".join(issues[:60]) or "(none)"

    prompt = (
        _ALIGNMENT_PROMPT_TEMPLATE
        .replace("LANG_PLACEHOLDER",        _lang_instruction(lang))
        .replace("TITLE_PLACEHOLDER",       spine.title)
        .replace("DISCIPLINE_PLACEHOLDER",  spine.discipline)
        .replace("PROBLEM_PLACEHOLDER",     spine.problem_statement[:600])
        .replace("GAP_PLACEHOLDER",         spine.research_gap[:400])
        .replace("RQS_PLACEHOLDER",         "; ".join(spine.research_questions) or "NOT FOUND")
        .replace("ROS_PLACEHOLDER",         "; ".join(spine.research_objectives) or "NOT FOUND")
        .replace("HYPS_PLACEHOLDER",        "; ".join(spine.hypotheses) or "NOT FOUND")
        .replace("THEORY_PLACEHOLDER",      spine.theory_used)
        .replace("VARS_PLACEHOLDER",        "; ".join(spine.variables) or "NOT FOUND")
        .replace("METHOD_PLACEHOLDER",      spine.methodology)
        .replace("SAMPLING_PLACEHOLDER",    spine.sampling)
        .replace("INSTRUMENT_PLACEHOLDER",  spine.instrument)
        .replace("ANALYSIS_PLACEHOLDER",    spine.analysis_technique)
        .replace("FINDINGS_PLACEHOLDER",    "; ".join(spine.key_findings) or "NOT FOUND")
        .replace("CONCLUSIONS_PLACEHOLDER", "; ".join(spine.conclusions) or "NOT FOUND")
        .replace("DIGESTS_PLACEHOLDER",     digests_block(chs))
        .replace("CRITICAL_ISSUES_PLACEHOLDER", issue_block)
    )
    try:
        d = _parse_json_loose(await _claude_text(_ALIGNMENT_SYSTEM, prompt, "alignment_audit",
                                                 schema=_ALIGNMENT_SCHEMA, document=thesis_document(chs)))
        rows = [AlignmentRow(**{k: _flatten_to_str(r.get(k, "—")) for k in (
                    "rq", "ro", "hypothesis", "theory_construct", "method", "instrument",
                    "analysis", "finding", "discussion", "conclusion", "broken_link",
                    "evidence", "note", "fix")},
                    status=str(r.get("status", "UNCLEAR")).upper())
                for r in d.get("rows", []) if isinstance(r, dict)]
        matrix = AlignmentMatrix(
            rows=rows,
            overall_verdict=            _flatten_to_str(d.get("overall_verdict","")),
            golden_thread_score=        str(d.get("golden_thread_score","UNCLEAR")).upper(),
            critical_gaps=              [_flatten_to_str(g) for g in _flatten_to_list(d.get("critical_gaps"))],
            structural_recommendations= [_flatten_to_str(g) for g in _flatten_to_list(d.get("structural_recommendations"))],
            chapter_links=              [x for x in _flatten_to_list(d.get("chapter_links")) if isinstance(x, dict)],
            construct_consistency=      [x for x in _flatten_to_list(d.get("construct_consistency")) if isinstance(x, dict)],
        )
        print(f"[ThesisSifu] Alignment audit: {len(rows)} rows, {len(matrix.chapter_links)} links, "
              f"{len(matrix.construct_consistency)} constructs, score={matrix.golden_thread_score}")
        return matrix
    except json.JSONDecodeError as e:
        print(f"[ThesisSifu] Alignment JSON parse error: {e}")
        return AlignmentMatrix(overall_verdict="Alignment JSON parse failed.", golden_thread_score="UNCLEAR")
    except Exception as e:
        print(f"[ThesisSifu] Alignment audit error: {e}\n{traceback.format_exc()}")
        return AlignmentMatrix(overall_verdict=f"Alignment audit failed: {e}", golden_thread_score="UNCLEAR")


# ── Stage 4: Holistic Examiner (Claude) ───────────────────────
_EXAMINER_SYSTEM = """\
Write the examiner report on the thesis provided above. Read the thesis itself: the spine,
alignment audit and issue summary below are aids prepared by assistants — confirm each point
against the text before using it, and drop any that the text contradicts. Never judge a chapter
by how many comments it received. The spine was extracted automatically and may be incomplete;
"NOT FOUND" means the extraction missed it, not that the thesis lacks it.

For every weakness you raise:
  1. Locate it (chapter / section) and state the evidence (short quote or precise paraphrase).
  2. Explain why an examiner would object — which part of the golden thread or validity it harms.
  3. Give a concrete fix: what to add, remove or restructure, with an example sentence, table
     structure or analysis step where useful.
  4. Support it with literature: cite ONLY works from VERIFIED LITERATURE or the CANONICAL
     LIBRARIES (author, year). Never invent a reference, title or DOI.
Acknowledge genuine strengths briefly where they are earned.
Keep every section heading on its own line in the exact form "SECTION N — TITLE".\
"""

_EXAMINER_LANG_MS = """\
Write the whole report in formal academic Bahasa Melayu (DBP spelling and istilah). Keep the
word "SECTION" and its number in each heading but translate the title, e.g.
"SECTION 2 — ANALISIS BENANG EMAS". In the verdict section, state the outcome using the exact
English label listed, followed by its Bahasa Melayu translation in brackets. In the language
and tone section, assess formal academic Bahasa Melayu: DBP spelling and terminology,
consistent rendering of English technical terms (English term in brackets at first use), and
avoidance of code-mixing and colloquial language.\
"""

_PHD_EXAMINER = """\
External examiner for a PhD viva.

THESIS SPINE
============
Title:            "TITLE_PLACEHOLDER"
Field:            FIELD_PLACEHOLDER | Institution: INST_PLACEHOLDER
Problem:          PROBLEM_PLACEHOLDER
Research Qs:      RQS_PLACEHOLDER
Research Objs:    ROS_PLACEHOLDER
Theory:           THEORY_PLACEHOLDER
Methodology:      METHOD_PLACEHOLDER
Analysis:         ANALYSIS_PLACEHOLDER

ALIGNMENT VERDICT
=================
Golden thread score: GT_SCORE_PLACEHOLDER
Overall:             ALIGN_VERDICT_PLACEHOLDER
Critical gaps:
GAPS_PLACEHOLDER

CHAPTER + SUBSECTION ISSUE SUMMARY
===================================
SUMMARIES_PLACEHOLDER

CANONICAL LIBRARIES (recommend FROM these — do NOT invent references):
FRAMEWORKS_PLACEHOLDER
METHODS_PLACEHOLDER

Write a full examiner report with EXACTLY these sections:

SECTION 1 — ORIGINAL CONTRIBUTION TO KNOWLEDGE
SECTION 2 — GOLDEN THREAD ANALYSIS
Use the alignment audit detail. Trace each research question from Chapter 1 to the conclusion;
for each broken or partial link name the chapters involved, the evidence, and the fix.
SECTION 3 — CHAPTER + SUBSECTION ALIGNMENT AUDIT
For each chapter, list which subsections under-delivered on expected purpose.
SECTION 4 — THEORETICAL FRAMEWORK COHERENCE
Name a specific framework (with seminal author) to add or strengthen.
SECTION 5 — METHODOLOGICAL RIGOUR
Recommend specific named methods (e.g., PLS-SEM via Hair et al. 2017).
SECTION 6 — DATA QUALITY & ANALYTICAL DEPTH
SECTION 7 — SCOPUS-LEVEL LANGUAGE & TONE
SECTION 8 — CITATION & REFERENCE INTEGRITY
SECTION 9 — CRITICAL VIVA QUESTIONS
List 8 specific questions tied to spine and alignment gaps.
SECTION 10 — EXAMINER'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS REQUIRED / REFER (RESUBMIT) / FAIL
Formal statement 6-8 sentences.\
"""

_MASTERS_EXAMINER = """\
Internal reader for a Master's thesis viva.

THESIS SPINE
============
Title:         "TITLE_PLACEHOLDER"
Field:         FIELD_PLACEHOLDER | Institution: INST_PLACEHOLDER
Problem:       PROBLEM_PLACEHOLDER
Research Qs:   RQS_PLACEHOLDER
Research Objs: ROS_PLACEHOLDER
Theory:        THEORY_PLACEHOLDER
Methodology:   METHOD_PLACEHOLDER
Analysis:      ANALYSIS_PLACEHOLDER

ALIGNMENT VERDICT
=================
Golden thread score: GT_SCORE_PLACEHOLDER
Overall:             ALIGN_VERDICT_PLACEHOLDER
Critical gaps:
GAPS_PLACEHOLDER

CHAPTER + SUBSECTION ISSUE SUMMARY
====================================
SUMMARIES_PLACEHOLDER

CANONICAL LIBRARIES:
FRAMEWORKS_PLACEHOLDER
METHODS_PLACEHOLDER

Write a report with EXACTLY these sections:

SECTION 1 — THESIS ALIGNMENT & GOLDEN THREAD
Use the alignment audit detail: for each broken or partial link name the chapters involved,
the evidence, and the fix. Trace each research question from Chapter 1 to the conclusion.
SECTION 2 — CHAPTER + SUBSECTION ALIGNMENT AUDIT
SECTION 3 — THEORETICAL FRAMEWORK (recommend named theory + seminal author)
SECTION 4 — METHODOLOGICAL RIGOUR (recommend named method + seminal author)
SECTION 5 — DATA ROBUSTNESS & INSTRUMENTATION
SECTION 6 — SCOPUS-LEVEL LANGUAGE & TONE
SECTION 7 — CITATION INTEGRITY
SECTION 8 — CRITICAL VIVA QUESTIONS (5 questions tied to alignment gaps)
SECTION 9 — EXAMINER'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS REQUIRED / RESUBMIT
Top 5 issues that must be resolved.\
"""

_JOURNAL_EXAMINER = """\
Peer reviewer for a Scopus journal.

ARTICLE SPINE
=============
Title:       "TITLE_PLACEHOLDER"
Field:       FIELD_PLACEHOLDER
Problem:     PROBLEM_PLACEHOLDER
Research Qs: RQS_PLACEHOLDER
Theory:      THEORY_PLACEHOLDER
Methodology: METHOD_PLACEHOLDER
Analysis:    ANALYSIS_PLACEHOLDER

ALIGNMENT VERDICT
=================
Golden thread score: GT_SCORE_PLACEHOLDER
Overall:             ALIGN_VERDICT_PLACEHOLDER
Critical gaps:
GAPS_PLACEHOLDER

SECTION ISSUE SUMMARY
=====================
SUMMARIES_PLACEHOLDER

CANONICAL LIBRARIES:
FRAMEWORKS_PLACEHOLDER
METHODS_PLACEHOLDER

Write a review with EXACTLY these sections:

SECTION 1 — CONTRIBUTION & NOVELTY
SECTION 2 — LITERATURE CURRENCY & GAPS
SECTION 3 — METHODOLOGY
SECTION 4 — RESULTS & ANALYSIS
SECTION 5 — ARGUMENT COHERENCE
SECTION 6 — SCOPUS-LEVEL LANGUAGE & TONE
SECTION 7 — CITATION INTEGRITY
SECTION 8 — PUBLICATION VERDICT
One of: ACCEPT / MINOR REVISIONS / MAJOR REVISIONS / REJECT
Top 3 critical revisions.\
"""

_UG_EXAMINER = """\
Supervisor reviewing an undergraduate FYP.

PROJECT SPINE
=============
Title:       "TITLE_PLACEHOLDER"
Field:       FIELD_PLACEHOLDER
Problem:     PROBLEM_PLACEHOLDER
Research Qs: RQS_PLACEHOLDER
Methodology: METHOD_PLACEHOLDER
Analysis:    ANALYSIS_PLACEHOLDER

ALIGNMENT VERDICT
=================
Golden thread score: GT_SCORE_PLACEHOLDER
Overall:             ALIGN_VERDICT_PLACEHOLDER

SECTION ISSUE SUMMARY
=====================
SUMMARIES_PLACEHOLDER

CANONICAL LIBRARIES:
FRAMEWORKS_PLACEHOLDER
METHODS_PLACEHOLDER

Write a report with EXACTLY these sections:

SECTION 1 — SCOPE & RESEARCH QUESTION
SECTION 2 — LITERATURE REVIEW
SECTION 3 — METHODOLOGY
SECTION 4 — ANALYSIS & FINDINGS
SECTION 5 — CRITICAL THINKING
SECTION 6 — WRITING & ACADEMIC CONVENTIONS
SECTION 7 — TOP ISSUES TO FIX
SECTION 8 — SUPERVISOR'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS / FAIL\
"""

_EXAMINER_TEMPLATES = {
    "PHD": _PHD_EXAMINER, "MASTERS": _MASTERS_EXAMINER,
    "JOURNAL_ARTICLE": _JOURNAL_EXAMINER, "UNDERGRADUATE": _UG_EXAMINER,
}

async def run_examiner(doc_type: str, spine: ThesisSpine, field: str, inst: str,
                       chs: list[ChapterSummary], align: AlignmentMatrix,
                       verified_refs: list, lang: str) -> str:
    if not claude_client:
        return "Examiner synthesis unavailable — ANTHROPIC_API_KEY not set."
    # Per-chapter issue summary: critical + moderate, with their golden-thread links
    summaries = ""
    for cs in chs:
        n_c = sum(1 for c in cs.comments if c.severity=="CRITICAL")
        n_m = sum(1 for c in cs.comments if c.severity=="MODERATE")
        if not cs.comments:
            continue
        summaries += (f"\nCh.{cs.chapter_num} — {cs.chapter_title}: "
                      f"{len(cs.comments)} comments (C:{n_c} M:{n_m})\n")
        for sub in cs.subsections:
            top = ([c for c in sub.comments if c.severity=="CRITICAL"][:2] +
                   [c for c in sub.comments if c.severity=="MODERATE"][:1])
            if top:
                summaries += f"  §{sub.subsection_num} {sub.title[:50]}\n"
                for c in top:
                    link = f" [{c.golden_thread_link}]" if c.golden_thread_link else ""
                    summaries += f"    • ({c.severity}){link} {_shorten(c.issue, 220)}\n"
    summaries = summaries[:12000]

    gaps = "\n".join(f"  * {g[:300]}" for g in align.critical_gaps[:8]) or "  (none)"
    template = _EXAMINER_TEMPLATES.get(doc_type, _MASTERS_EXAMINER)

    align_detail = []
    for r in align.rows:
        align_detail.append(f"  - [{r.status}] {_shorten(r.rq, 140)}"
                            + (f" | broken: {_shorten(r.broken_link, 160)}" if r.broken_link else "")
                            + (f" | evidence: {_shorten(r.evidence, 200)}" if r.evidence else ""))
    for x in align.chapter_links:
        if str(x.get("status", "")).upper() != "ALIGNED":
            align_detail.append(f"  - [{x.get('status')}] {x.get('from_chapter')} -> {x.get('to_chapter')}: "
                                f"{_shorten(str(x.get('link','')), 120)} | {_shorten(str(x.get('evidence','')), 200)}")
    for x in align.construct_consistency:
        if str(x.get("status", "")).upper() != "ALIGNED":
            align_detail.append(f"  - [{x.get('status')}] construct '{x.get('construct')}': "
                                f"{_shorten(str(x.get('trace','')), 200)}")
    refs = "\n".join(f"  - {r}" for r in verified_refs[:40]) or "  (none retrieved)"
    extra = ("ALIGNMENT AUDIT DETAIL (from the cross-chapter golden-thread audit)\n"
             "===================================================================\n"
             + ("\n".join(align_detail)[:8000] or "  (no detail)") +
             "\n\nVERIFIED LITERATURE (real, retrieved records — the only new works you may cite)\n"
             "================================================================================\n"
             + refs + "\n\n" + _lang_instruction(lang)
             + ("\n" + _EXAMINER_LANG_MS if lang == "ms" else ""))
    prompt = (
        template
        .replace("TITLE_PLACEHOLDER",        spine.title[:150])
        .replace("FIELD_PLACEHOLDER",         field)
        .replace("INST_PLACEHOLDER",          inst)
        .replace("PROBLEM_PLACEHOLDER",       spine.problem_statement[:300])
        .replace("RQS_PLACEHOLDER",           "; ".join(spine.research_questions[:4]) or "NOT FOUND")
        .replace("ROS_PLACEHOLDER",           "; ".join(spine.research_objectives[:4]) or "NOT FOUND")
        .replace("THEORY_PLACEHOLDER",        spine.theory_used[:100])
        .replace("METHOD_PLACEHOLDER",        spine.methodology[:100])
        .replace("ANALYSIS_PLACEHOLDER",      spine.analysis_technique[:100])
        .replace("GT_SCORE_PLACEHOLDER",      align.golden_thread_score)
        .replace("ALIGN_VERDICT_PLACEHOLDER", align.overall_verdict[:1200])
        .replace("GAPS_PLACEHOLDER",          gaps)
        .replace("SUMMARIES_PLACEHOLDER",     summaries)
        .replace("FRAMEWORKS_PLACEHOLDER",    CANONICAL_FRAMEWORKS)
        .replace("METHODS_PLACEHOLDER",       CANONICAL_METHODS)
    )
    # Every template ends with its "Write a ... with EXACTLY these sections" instruction;
    # the alignment detail and verified literature go just before it.
    prompt = prompt.replace("\nWrite a", "\n" + extra + "\n\nWrite a", 1)
    try:
        return await _claude_text(_EXAMINER_SYSTEM, prompt, "examiner_report",
                                  document=thesis_document(chs))
    except Exception as e:
        print(f"[ThesisSifu] Examiner error: {e}\n{traceback.format_exc()}")
        return f"Examiner synthesis failed: {e}"


# ── PDF helpers ────────────────────────────────────────────────
def pdf_styles():
    def S(n, **k):
        k.setdefault("fontName","Times-Roman"); k.setdefault("fontSize",11)
        k.setdefault("leading",16); k.setdefault("textColor",BLACK)
        k.setdefault("alignment",TA_JUSTIFY)
        return ParagraphStyle(n,**k)
    return {
        "Title":   S("Title",fontName="Helvetica-Bold",fontSize=20,leading=26,textColor=NAVY,alignment=TA_LEFT),
        "Sub":     S("Sub",fontName="Helvetica",fontSize=12,leading=18,textColor=ACCENT,alignment=TA_LEFT),
        "Meta":    S("Meta",fontName="Helvetica",fontSize=10,leading=15,textColor=colors.HexColor("#444444"),alignment=TA_LEFT),
        "Badge":   S("Badge",fontName="Helvetica-Bold",fontSize=10,leading=14,textColor=WHITE,alignment=TA_CENTER),
        "SecHead": S("SecHead",fontName="Helvetica-Bold",fontSize=13,leading=20,textColor=NAVY,alignment=TA_LEFT,spaceBefore=16,spaceAfter=6),
        "ChHead":  S("ChHead",fontName="Helvetica-Bold",fontSize=11,leading=16,textColor=ACCENT,alignment=TA_LEFT,spaceBefore=12,spaceAfter=4),
        "SubHead": S("SubHead",fontName="Helvetica-Bold",fontSize=10,leading=14,textColor=colors.HexColor("#2A5A8A"),alignment=TA_LEFT,spaceBefore=8,spaceAfter=3),
        "Body":    S("Body",leading=18,spaceAfter=6),
        "Bullet":  S("Bullet",leading=17,leftIndent=16,spaceAfter=4),
        "Verdict": S("Verdict",fontName="Helvetica-Bold",fontSize=15,leading=22,textColor=WHITE,alignment=TA_CENTER),
        "VrdSub":  S("VrdSub",fontName="Helvetica",fontSize=10,leading=14,textColor=WHITE,alignment=TA_CENTER),
        "TblH":    S("TblH",fontName="Helvetica-Bold",fontSize=9,leading=13,textColor=NAVY,alignment=TA_LEFT),
        "TblC":    S("TblC",fontName="Times-Roman",fontSize=9,leading=13,textColor=BLACK,alignment=TA_LEFT),
        "TblCSm":  S("TblCSm",fontName="Times-Roman",fontSize=8,leading=11,textColor=BLACK,alignment=TA_LEFT),
        "Footer":  S("Footer",fontName="Helvetica-Oblique",fontSize=8,leading=12,textColor=colors.HexColor("#999"),alignment=TA_CENTER),
        "Excerpt": S("Excerpt",fontName="Times-Italic",fontSize=9,leading=13,textColor=colors.HexColor("#555"),leftIndent=8),
        "Detail":  S("Detail",fontName="Times-Roman",fontSize=10,leading=15,textColor=BLACK,leftIndent=4),
        "SpineV":  S("SpineV",fontName="Times-Roman",fontSize=10,leading=14,textColor=BLACK),
        "SpineK":  S("SpineK",fontName="Helvetica-Bold",fontSize=10,leading=14,textColor=NAVY),
    }

def HR(c=RULE,t=0.5,b=5,a=7):
    return HRFlowable(width="100%",thickness=t,color=c,spaceBefore=b,spaceAfter=a)

def _tbl(data, col_widths, styles):
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle(styles))
    return t

def meta_box(S, rows):
    data = [[Paragraph(f'<font name="Helvetica-Bold">{k}</font>',S["Meta"]),
             Paragraph(str(v),S["Meta"])] for k,v in rows]
    return _tbl(data,[4*cm,W-5*cm-4*cm],[
        ("BACKGROUND",(0,0),(-1,-1),BOX_BG),("BOX",(0,0),(-1,-1),.5,RULE),
        ("INNERGRID",(0,0),(-1,-1),.3,RULE),("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),("LEFTPADDING",(0,0),(-1,-1),8),
        ("RIGHTPADDING",(0,0),(-1,-1),8),("VALIGN",(0,0),(-1,-1),"TOP")])

def vbanner(S, text, color, subtitle="Examiner's Recommended Outcome"):
    t = Table([[Paragraph(text.upper(),S["Verdict"])],[Paragraph(subtitle,S["VrdSub"])]],
              colWidths=[W-5*cm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),color),
        ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("LEFTPADDING",(0,0),(-1,-1),16),("RIGHTPADDING",(0,0),(-1,-1),16)]))
    return t

def sev_badge(S, sev, count, lang="en"):
    t = Table([[Paragraph(f"{sev_label(sev, lang)}: {count}",S["Badge"])]],colWidths=[3.5*cm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),SEV_COLORS.get(sev,ACCENT)),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8)]))
    return t

def type_badge_table(S, text, color):
    t = Table([[Paragraph(text,S["Badge"])]],colWidths=[W-5*cm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),color),
        ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
        ("LEFTPADDING",(0,0),(-1,-1),12),("RIGHTPADDING",(0,0),(-1,-1),12)]))
    return t

_SPINE_KEYS = [
    ("Title", "Tajuk"), ("Discipline", "Bidang"), ("Problem Statement", "Pernyataan Masalah"),
    ("Research Gap", "Jurang Kajian"), ("Research Questions", "Persoalan Kajian"),
    ("Research Objectives", "Objektif Kajian"), ("Theory Used", "Teori"),
    ("Variables", "Pemboleh Ubah"), ("Methodology", "Metodologi"), ("Sampling", "Persampelan"),
    ("Instrument", "Instrumen"), ("Analysis Technique", "Teknik Analisis"),
    ("Key Findings", "Dapatan Utama"), ("Conclusions", "Kesimpulan"),
]

def spine_box(S, spine: ThesisSpine, lang: str = "en"):
    vals = [
        spine.title[:120], spine.discipline, spine.problem_statement[:300], spine.research_gap[:250],
        "; ".join(spine.research_questions[:6]) or "NOT FOUND",
        "; ".join(spine.research_objectives[:6]) or "NOT FOUND",
        spine.theory_used, "; ".join(spine.variables[:10]) or "NOT FOUND",
        spine.methodology, spine.sampling, spine.instrument, spine.analysis_technique,
        "; ".join(spine.key_findings[:4]) or "NOT FOUND",
        "; ".join(spine.conclusions[:4]) or "NOT FOUND",
    ]
    rows = [(ms if lang == "ms" else en, v) for (en, ms), v in zip(_SPINE_KEYS, vals)]
    data = [[Paragraph(k,S["SpineK"]), _safe_para(str(v),S["SpineV"])] for k,v in rows]
    return _tbl(data,[4.5*cm,W-5*cm-4.5*cm],[
        ("BACKGROUND",(0,0),(-1,-1),BOX_BG),("BOX",(0,0),(-1,-1),.5,RULE),
        ("INNERGRID",(0,0),(-1,-1),.3,RULE),("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),8),
        ("RIGHTPADDING",(0,0),(-1,-1),8),("VALIGN",(0,0),(-1,-1),"TOP")])

def parse_secs(text):
    pat = re.compile(r'(?:^|\n)\s*(?:#{1,6}\s*)?(?:\*{0,2})?(?:SECTION|SEKSYEN|BAHAGIAN)\s+\d+\s*[—–\-:]+\s*([^\n*]+?)(?:\*{0,2})?\s*\n',re.IGNORECASE)
    ms  = list(pat.finditer(text))
    if not ms: return [{"title":"Full Report","body":text.strip()}]
    out = []
    for i,m in enumerate(ms):
        t = m.group(1).strip().rstrip("*").strip()
        s = m.end(); e = ms[i+1].start() if i+1<len(ms) else len(text)
        out.append({"title":t,"body":text[s:e].strip()})
    return out

_VERDICTS = [
    ("READY FOR SUBMISSION", GREEN), ("PASS WITH MINOR CORRECTIONS", AMBER),
    ("MINOR REVISIONS REQUIRED", AMBER), ("MINOR REVISIONS", AMBER),
    ("MAJOR REVISIONS REQUIRED", RED), ("MAJOR REVISIONS", RED),
    ("REFER (RESUBMIT)", RED), ("REJECT AND REWRITE", RED), ("REJECT", RED),
    ("RESUBMIT", RED), ("FAIL", RED), ("ACCEPT", GREEN), ("PASS", GREEN),
]

def detect_verdict(text):
    """Find the verdict label in the verdict (last) section only, as a whole word.
    Searching the whole report matched words like "acceptable" or "pass" in
    ordinary sentences and showed the wrong banner."""
    secs = parse_secs(text)
    scope = secs[-1]["body"] if secs else text
    u = scope.upper()
    for lbl, col in _VERDICTS:
        if re.search(r'(?<![A-Z])' + re.escape(lbl) + r'(?![A-Z])', u):
            return lbl.title(), col
    return "Review Complete", ACCENT

def _safe_para(text: str, style) -> "Paragraph":
    """Create a ReportLab Paragraph, stripping any HTML tags that would cause
    XML parse errors (e.g. mismatched <b>/<i> from Claude's markdown output)."""
    import html as html_mod
    # 1. Escape raw ampersands that aren't already entities
    text = re.sub(r'&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[\da-fA-F]+);)', '&amp;', text)
    # 2. Strip any XML/HTML angle-bracket tags that are malformed or unknown to ReportLab
    #    Keep only the tags ReportLab actually supports: b, i, u, br, font, super, sub
    allowed = {'b','i','u','br','font','super','sub','strike'}
    def _strip_bad_tags(m):
        tag = m.group(1).lower().split()[0].lstrip('/')
        return m.group(0) if tag in allowed else ''
    text = re.sub(r'<(/?\w[^>]*)>', _strip_bad_tags, text)
    # 3. Balance remaining b/i tags — close any that were opened but not closed
    for tag in ('b', 'i', 'u'):
        opens  = len(re.findall(f'<{tag}>', text))
        closes = len(re.findall(f'</{tag}>', text))
        if opens > closes:
            text += f'</{tag}>' * (opens - closes)
        elif closes > opens:
            text = f'<{tag}>' * (closes - opens) + text
    try:
        return Paragraph(text, style)
    except Exception:
        # Last resort: strip ALL tags and render as plain text
        plain = re.sub(r'<[^>]+>', '', text)
        return Paragraph(plain, style)

def body2story(S, body):
    items = []
    for line in body.split("\n"):
        line = line.strip()
        if not line:
            items.append(Spacer(1,3))
            continue
        # Convert markdown bold/italic to HTML — handle ***text*** (bold+italic) first
        line = re.sub(r'\*\*\*(.+?)\*\*\*', r'<b><i>\1</i></b>', line)
        line = re.sub(r'\*\*(.+?)\*\*',     r'<b>\1</b>',         line)
        line = re.sub(r'\*(.+?)\*',          r'<i>\1</i>',         line)
        # Strip markdown headers (### ## #) — ReportLab can't render them
        line = re.sub(r'^#{1,6}\s+', '', line)
        if re.match(r'^[-•]\s+', line):
            items.append(_safe_para("• " + line[2:], S["Bullet"]))
        elif re.match(r'^\d+[\.\)]\s+', line):
            items.append(_safe_para(line, S["Bullet"]))
        else:
            items.append(_safe_para(line, S["Body"]))
    return items

def _comment_text(c: ParagraphComment, lang: str = "en") -> str:
    sev = f"[{sev_label(c.severity, lang)}]"
    up  = lambda k: L(k, lang).upper()
    gt  = f"\n\n{up('thread_link')}: {c.golden_thread_link}" if c.golden_thread_link else ""
    fw  = f"\n\n{up('framework')}: {c.suggested_framework}" if c.suggested_framework else ""
    mt  = f"\n\n{up('method')}: {c.suggested_method}" if c.suggested_method else ""
    refs = ""
    if c.references:
        refs = f"\n\n{up('verified_refs')}:" + "".join(
            f"\n• {r['citation']}" + (f" — {r['how_to_use']}" if r.get("how_to_use") else "")
            for r in c.references)
    return (f"{sev} ThesisSifu Pro — §{c.subsection} {c.subsection_title}\n\n"
            f"{up('issue')}: {c.issue}{gt}\n\n"
            f"{up('recommendation')}: {c.recommendation}\n\n"
            f"{up('literature')}: {c.literature_needed}\n\n"
            f"{up('theory')}: {c.theory_needed}"
            f"{fw}{mt}{refs}")


# ── Output 1: Examiner Report PDF ─────────────────────────────
def build_examiner_pdf(filename, audit_id, doc_type, clf, spine, examiner_text, align, chs, lang="en") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf,pagesize=A4,leftMargin=2.5*cm,rightMargin=2.5*cm,
                            topMargin=2.2*cm,bottomMargin=2.2*cm)
    S = pdf_styles(); story = []

    type_label = TYPE_LABELS.get(doc_type,doc_type)
    story.append(type_badge_table(S,type_label.upper(),TYPE_COLORS.get(doc_type,NAVY)))
    story.append(Spacer(1,14))
    story.append(Paragraph(L("examiner_title", lang),S["Title"]))
    story.append(Paragraph("ThesisSifu Pro v4.2 — Alignment-Aware Multi-Agent Panel",S["Sub"]))
    story.append(HR(c=NAVY,t=1.5,b=4,a=10))

    au=clf.get("authors","UNKNOWN"); fi=clf.get("field","UNKNOWN"); ins=clf.get("institution","UNKNOWN")
    story.append(meta_box(S,[
        ("Document",filename),("Detected As",type_label),
        ("Confidence",clf.get("confidence","N/A")),
        ("Author(s)",(au[:80]+"…") if len(au)>80 else au),
        ("Field",fi),("Institution",ins),
        ("Audit ID",audit_id),("Date",datetime.now().strftime("%d %B %Y, %H:%M")),
        ("Report","1 of 4 — Examiner Audit Report"),
    ]))
    story.append(Spacer(1,10))

    story.append(Paragraph(L("spine", lang),S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))
    story.append(spine_box(S, spine, lang))
    story.append(Spacer(1,10))

    nc=sum(1 for cs in chs for c in cs.comments if c.severity=="CRITICAL")
    nm=sum(1 for cs in chs for c in cs.comments if c.severity=="MODERATE")
    ns=sum(1 for cs in chs for c in cs.comments if c.severity=="SUGGESTION")
    brow = Table([[sev_badge(S,"CRITICAL",nc,lang),Spacer(6,1),
                   sev_badge(S,"MODERATE",nm,lang),Spacer(6,1),
                   sev_badge(S,"SUGGESTION",ns,lang)]],
                 colWidths=[3.5*cm,.5*cm,3.5*cm,.5*cm,3.5*cm])
    brow.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))
    story.append(brow); story.append(Spacer(1,10))

    gt = align.golden_thread_score
    gt_color = {"STRONG":GREEN,"ACCEPTABLE":AMBER,"WEAK":RED,"BROKEN":RED}.get(gt, ACCENT)
    story.append(vbanner(S,f"{L('golden_thread', lang)}: {gt}",gt_color,subtitle=L("align_score_sub", lang)))
    story.append(Spacer(1,8))
    vl,vc = detect_verdict(examiner_text)
    story.append(vbanner(S,vl,vc,subtitle=L("verdict_sub", lang))); story.append(Spacer(1,10)); story.append(HR())

    for sec in parse_secs(examiner_text):
        bl = [Paragraph(sec["title"].upper(),S["SecHead"]),HR(c=ACCENT,t=.7,b=0,a=6)]
        bl.extend(body2story(S,sec["body"])); bl.append(Spacer(1,6))
        story.append(KeepTogether(bl[:3])); story.extend(bl[3:])

    story.append(PageBreak())
    story.append(Paragraph(L("issue_summary", lang),S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=8))
    td = [[Paragraph(h,S["TblH"]) for h in ["Chapter / Subsection","Critical","Moderate","Suggestions","Total"]]]
    for cs in chs:
        nc2=sum(1 for c in cs.comments if c.severity=="CRITICAL")
        nm2=sum(1 for c in cs.comments if c.severity=="MODERATE")
        ns2=len(cs.comments)-nc2-nm2
        td.append([Paragraph(f"<b>Ch.{cs.chapter_num} — {_shorten(cs.chapter_title, 40)}</b>",S["TblC"]),
                   Paragraph(str(nc2),S["TblC"]),Paragraph(str(nm2),S["TblC"]),
                   Paragraph(str(ns2),S["TblC"]),Paragraph(str(nc2+nm2+ns2),S["TblC"])])
        for sub in cs.subsections:
            if not sub.comments: continue
            snc=sum(1 for c in sub.comments if c.severity=="CRITICAL")
            snm=sum(1 for c in sub.comments if c.severity=="MODERATE")
            sns=len(sub.comments)-snc-snm
            td.append([Paragraph(f"   §{sub.subsection_num} {_shorten(sub.title, 45)}",S["TblCSm"]),
                       Paragraph(str(snc),S["TblCSm"]),Paragraph(str(snm),S["TblCSm"]),
                       Paragraph(str(sns),S["TblCSm"]),Paragraph(str(snc+snm+sns),S["TblCSm"])])
    cw = W-5*cm
    t = Table(td,colWidths=[cw*.52,cw*.12,cw*.12,cw*.12,cw*.12], repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),("BOX",(0,0),(-1,-1),.5,RULE),
        ("INNERGRID",(0,0),(-1,-1),.3,RULE),("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),5),
        ("RIGHTPADDING",(0,0),(-1,-1),5),("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(t)
    story.append(Spacer(1,20)); story.append(HR())
    story.append(Paragraph("ThesisSifu Pro v4.2 — See Report 2 (Annotated Thesis), "
        "Report 3 (Commentary Log), Report 4 (Alignment Matrix) in this ZIP.",S["Footer"]))
    doc.build(story); buf.seek(0); return buf.getvalue()


# ── Output 2: Annotated DOCX ──────────────────────────────────
def _get_or_create_comments_part(doc):
    rt = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
    for rel in doc.part.rels.values():
        if rel.reltype == rt: return rel.target_part
    xml_str = '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    element = parse_xml(xml_str)
    uri = PackURI('/word/comments.xml')
    ct  = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    cp  = XmlPart(uri, ct, element, doc.part.package)
    doc.part.relate_to(cp, rt)
    return cp

def _insert_comment(doc, para, cid, author, text, date_str):
    cp = _get_or_create_comments_part(doc)
    ce = OxmlElement('w:comment')
    ce.set(qn('w:id'),str(cid)); ce.set(qn('w:author'),author)
    ce.set(qn('w:date'),date_str); ce.set(qn('w:initials'),'TS')
    pe = OxmlElement('w:p'); re_ = OxmlElement('w:r'); te = OxmlElement('w:t')
    te.text = text; te.set(qn('xml:space'),'preserve')
    re_.append(te); pe.append(re_); ce.append(pe)
    cp.element.append(ce)
    px = para._p
    crs = OxmlElement('w:commentRangeStart'); crs.set(qn('w:id'),str(cid)); px.insert(0,crs)
    cre = OxmlElement('w:commentRangeEnd');   cre.set(qn('w:id'),str(cid)); px.append(cre)
    rr  = OxmlElement('w:r'); rpr = OxmlElement('w:rPr')
    rs  = OxmlElement('w:rStyle'); rs.set(qn('w:val'),'CommentReference')
    rpr.append(rs); rr.append(rpr)
    cr  = OxmlElement('w:commentReference'); cr.set(qn('w:id'),str(cid)); rr.append(cr); px.append(rr)

def build_annotated_docx(content, filename, audit_id, chs, clf, lang="en") -> bytes:
    doc = DocxDocument(io.BytesIO(content))
    date_str = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    cp = _get_or_create_comments_part(doc)
    existing_ids = [int(c.get(qn('w:id'))) for c in cp.element.xpath('.//w:comment')
                    if c.get(qn('w:id')) is not None]
    cid = max(existing_ids)+1 if existing_ids else 0

    lookup: dict[str,list] = {}
    for cs in chs:
        for c in cs.comments:
            key = c.para_excerpt[:50].lower().strip()
            lookup.setdefault(key,[]).append(c)

    matched = set()
    for para in doc.paragraphs:
        pt = para.text.strip()
        if not pt or len(pt) < 30: continue
        pkey = pt[:50].lower().strip()
        best, best_sc = None, 0
        for ekey, cmts in lookup.items():
            aw = set(pkey.split()); bw = set(ekey.split())
            sc = len(aw & bw) / max(len(aw),1)
            if sc > best_sc and sc > 0.30:
                best_sc = sc; best = (ekey, cmts)
        if best:
            for c in best[1]:
                if id(c) in matched: continue
                matched.add(id(c))
                try:
                    _insert_comment(doc, para, cid, "ThesisSifu Pro", _comment_text(c, lang), date_str)
                    cid += 1
                except Exception as e:
                    print(f"[ThesisSifu] DOCX comment insert error {cid}: {e}")
                    cid += 1

    hdr = doc.sections[0].header
    if hdr.paragraphs:
        hdr.paragraphs[0].text = (f"ThesisSifu Pro v4.2 | Annotated Thesis | "
                                   f"Audit ID: {audit_id} | {datetime.now().strftime('%d %B %Y')}")
    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf.getvalue()


# ── Output 2 (PDF): Annotated PDF (Sticky Notes) ─────────────
def build_annotated_pdf(content: bytes, full_text: str, audit_id: str, chs: list, lang: str = "en") -> bytes:
    """Place sticky-note annotations on the PDF.

    PDF TEXT SEARCH LIMITATION: Many thesis PDFs use Identity-H CID TrueType fonts
    which PyMuPDF cannot search with page.search_for(). Rather than placing all
    annotations at (30,30) as invisible stacked fallbacks, we use a reliable
    position-based strategy:
      - Place each annotation in the RIGHT MARGIN of the estimated page
      - Stagger vertically so multiple annotations on one page don't overlap
      - Use a different Y offset for CRITICAL (top third), MODERATE (mid), SUGGESTION (bottom)
      - Colour-code icon: red=CRITICAL, orange=MODERATE, blue=SUGGESTION
    This guarantees every comment is visible and readable, regardless of font encoding.
    """
    fitz_doc = fitz.open(stream=content, filetype="pdf")
    total_pages = len(fitz_doc)

    # Track how many annotations are already on each page (to stagger Y)
    page_annot_counts: dict[int, int] = {}

    all_comments = [c for cs in chs for c in cs.comments]
    print(f"[ThesisSifu] Placing {len(all_comments)} PDF annotations (position-based)")

    for c in all_comments:
        txt = _comment_text(c, lang)

        # Determine target page (0-indexed, clamped)
        pg = max(0, min(c.page_estimate - 1, total_pages - 1))

        # Try text search first (works when fonts cooperate)
        search_text = c.para_excerpt[:40].replace('\n', ' ').strip()
        placed_by_search = False

        if search_text:
            search_range = range(max(0, pg - 2), min(total_pages, pg + 3))
            for p_num in search_range:
                page = fitz_doc[p_num]
                rects = page.search_for(search_text)
                if not rects and len(search_text) > 20:
                    rects = page.search_for(search_text[:20])
                if rects:
                    rect  = rects[0]
                    point = fitz.Point(max(5, rect.x0 - 18), rect.y0)
                    annot = page.add_text_annot(point, txt)
                    annot.set_info(title="ThesisSifu Pro", content=txt)
                    if   c.severity == "CRITICAL":   annot.set_colors(stroke=(0.8, 0, 0))
                    elif c.severity == "MODERATE":   annot.set_colors(stroke=(0.9, 0.4, 0))
                    else:                             annot.set_colors(stroke=(0, 0, 0.8))
                    annot.update()
                    placed_by_search = True
                    page_annot_counts[p_num] = page_annot_counts.get(p_num, 0) + 1
                    break

        if not placed_by_search:
            # ── Position-based right-margin placement ──────────────────────
            page = fitz_doc[pg]
            pw   = page.rect.width
            ph   = page.rect.height

            # Base Y by severity: spread across page to avoid clustering
            n = page_annot_counts.get(pg, 0)

            # CRITICAL → top third, MODERATE → mid, SUGGESTION → bottom third
            if c.severity == "CRITICAL":
                base_y = ph * 0.10
            elif c.severity == "MODERATE":
                base_y = ph * 0.40
            else:
                base_y = ph * 0.70

            # Stagger: each successive annotation on this page moves down 28pt
            # Wrap back toward base if we'd go off page
            y_pos = base_y + (n % 8) * 28
            if y_pos > ph - 30:
                y_pos = base_y + (n % 4) * 14

            # X: right margin (5pt from edge)
            x_pos = pw - 18

            point = fitz.Point(x_pos, y_pos)
            annot = page.add_text_annot(point, txt)
            annot.set_info(title="ThesisSifu Pro", content=txt)
            if   c.severity == "CRITICAL":   annot.set_colors(stroke=(0.8, 0, 0))
            elif c.severity == "MODERATE":   annot.set_colors(stroke=(0.9, 0.4, 0))
            else:                             annot.set_colors(stroke=(0, 0, 0.8))
            annot.update()
            page_annot_counts[pg] = page_annot_counts.get(pg, 0) + 1

    total_placed = sum(page_annot_counts.values())
    print(f"[ThesisSifu] PDF annotations placed: {total_placed} across "
          f"{len(page_annot_counts)} pages")
    return fitz_doc.tobytes()   # tobytes() is more reliable than write() for in-memory


# ── Output 3: Commentary Report PDF ───────────────────────────
def build_commentary_pdf(filename, audit_id, doc_type, clf, spine, chs, lang="en") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf,pagesize=A4,leftMargin=2.5*cm,rightMargin=2.5*cm,
                            topMargin=2.2*cm,bottomMargin=2.2*cm)
    S = pdf_styles(); story = []

    story.append(type_badge_table(S,L("commentary_badge", lang),TYPE_COLORS.get(doc_type,NAVY)))
    story.append(Spacer(1,14))
    story.append(Paragraph(L("commentary_title", lang),S["Title"]))
    story.append(Paragraph("ThesisSifu Pro v4.2 — " + L("commentary_sub", lang),S["Sub"]))
    story.append(HR(c=NAVY,t=1.5,b=4,a=10))

    au = clf.get("authors","UNKNOWN")
    story.append(meta_box(S,[
        ("Document",filename),("Type",TYPE_LABELS.get(doc_type,doc_type)),
        ("Title",(spine.title[:85]+"…") if len(spine.title)>85 else spine.title),
        ("Author(s)",(au[:80]+"…") if len(au)>80 else au),
        ("Audit ID",audit_id),("Date",datetime.now().strftime("%d %B %Y, %H:%M")),
        ("Report","3 of 4 — Paragraph Commentary Log"),
    ]))
    story.append(Spacer(1,10))

    all_c = [c for cs in chs for c in cs.comments]
    nc = sum(1 for c in all_c if c.severity=="CRITICAL")
    nm = sum(1 for c in all_c if c.severity=="MODERATE")
    ns = sum(1 for c in all_c if c.severity=="SUGGESTION")

    story.append(Paragraph(L("distribution", lang),S["SecHead"]))
    cw = W-5*cm
    td = [[Paragraph(h,S["TblH"]) for h in ["Severity","Count","Description"]],
          [Paragraph("CRITICAL",S["TblC"]),Paragraph(str(nc),S["TblC"]),
           Paragraph("Fundamental flaws requiring resolution",S["TblC"])],
          [Paragraph("MODERATE",S["TblC"]),Paragraph(str(nm),S["TblC"]),
           Paragraph("Significant weaknesses requiring revision",S["TblC"])],
          [Paragraph("SUGGESTION",S["TblC"]),Paragraph(str(ns),S["TblC"]),
           Paragraph("Improvement opportunities",S["TblC"])],
          [Paragraph("TOTAL",S["TblC"]),Paragraph(str(len(all_c)),S["TblC"]),Paragraph("",S["TblC"])]]
    t = Table(td,colWidths=[cw*.22,cw*.12,cw*.66], repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),("BOX",(0,0),(-1,-1),.5,RULE),
        ("INNERGRID",(0,0),(-1,-1),.3,RULE),("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),("LEFTPADDING",(0,0),(-1,-1),7),
        ("RIGHTPADDING",(0,0),(-1,-1),7)]))
    story.append(t); story.append(Spacer(1,16)); story.append(HR())

    if not all_c:
        story.append(Paragraph("<i>No paragraph comments were generated. "
            "Check Railway logs for API errors or empty document detection.</i>",S["Body"]))
    else:
        for cs in chs:
            if not cs.comments: continue
            story.append(Paragraph(f"CHAPTER {cs.chapter_num} — {cs.chapter_title.upper()}",S["ChHead"]))
            nc2=sum(1 for c in cs.comments if c.severity=="CRITICAL")
            nm2=sum(1 for c in cs.comments if c.severity=="MODERATE")
            story.append(Paragraph(f"<b>{len(cs.comments)} comments</b> — "
                f"Critical: {nc2} | Moderate: {nm2} | Suggestions: {len(cs.comments)-nc2-nm2}",S["Body"]))
            story.append(HR(c=RULE,t=.4,b=2,a=6))

            for sub in cs.subsections:
                if not sub.comments: continue
                story.append(Paragraph(f"§{sub.subsection_num} — {sub.title}",S["SubHead"]))
                story.append(_safe_para(f"<i>{L('expected', lang)}: {sub.expected_purpose}</i>",S["Excerpt"]))
                story.append(Spacer(1,4))

                for i, c in enumerate(sub.comments,1):
                    sc = SEV_COLORS.get(c.severity,ACCENT)
                    hdr_row = Table(
                        [[Paragraph(sev_label(c.severity, lang),S["Badge"]),Spacer(6,1),
                          Paragraph(f"§{sub.subsection_num} Comment {i} | Page ~{c.page_estimate}",S["Meta"])]],
                        colWidths=[2.2*cm,.4*cm,W-5*cm-2.6*cm])
                    hdr_row.setStyle(TableStyle([("BACKGROUND",(0,0),(0,0),sc),
                        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
                        ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                        ("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
                    # Sanitize excerpt — may contain quotes, ampersands, angle brackets
                    safe_excerpt = c.para_excerpt[:120].replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                    excerpt_box = Table(
                        [[_safe_para(f'<i>"{safe_excerpt}..."</i>',S["Excerpt"])]],
                        colWidths=[W-5*cm])
                    excerpt_box.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),LGRAY),
                        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
                        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8)]))
                    esc = lambda x: str(x).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                    gt_line = (f'<br/><br/><b>{L("thread_link", lang)}:</b> {esc(c.golden_thread_link)}'
                               if c.golden_thread_link else "")
                    fw_line = f'<br/><br/><b>{L("framework", lang)}:</b> {esc(c.suggested_framework)}' if c.suggested_framework else ""
                    mt_line = f'<br/><br/><b>{L("method", lang)}:</b> {esc(c.suggested_method)}' if c.suggested_method else ""
                    ref_line = ""
                    if c.references:
                        ref_line = f'<br/><br/><b>{L("verified_refs", lang)}:</b>' + "".join(
                            f'<br/>• {esc(r["citation"])}'
                            + (f' <i>— {esc(r["how_to_use"])}</i>' if r.get("how_to_use") else "")
                            for r in c.references)
                    detail_box = Table(
                        [[_safe_para(
                            f'<b>{L("issue", lang)}:</b> {esc(c.issue)}{gt_line}<br/><br/>'
                            f'<b>{L("recommendation", lang)}:</b> {esc(c.recommendation)}<br/><br/>'
                            f'<b>{L("literature", lang)}:</b> {esc(c.literature_needed)}<br/><br/>'
                            f'<b>{L("theory", lang)}:</b> {esc(c.theory_needed)}'
                            f'{fw_line}{mt_line}{ref_line}',S["Detail"])]],
                        colWidths=[W-5*cm])
                    detail_box.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),BOX_BG),
                        ("BOX",(0,0),(-1,-1),.4,RULE),("TOPPADDING",(0,0),(-1,-1),7),
                        ("BOTTOMPADDING",(0,0),(-1,-1),7),("LEFTPADDING",(0,0),(-1,-1),10),
                        ("RIGHTPADDING",(0,0),(-1,-1),10)]))
                    block = [hdr_row,excerpt_box,detail_box,Spacer(1,8)]
                    story.append(KeepTogether(block[:2])); story.extend(block[2:])
            story.append(Spacer(1,12)); story.append(HR())

    story.append(Spacer(1,6))
    story.append(Paragraph("ThesisSifu Pro v4.2 — Subsection commentary. "
        "See Report 1 (Examiner Audit), Report 2 (Annotated Thesis), "
        "Report 4 (Alignment Matrix).",S["Footer"]))
    doc.build(story); buf.seek(0); return buf.getvalue()


# ── Output 4: Alignment Matrix Report PDF ─────────────────────
def build_alignment_pdf(filename, audit_id, doc_type, clf, spine, align, chs, lang="en") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf,pagesize=A4,leftMargin=2*cm,rightMargin=2*cm,
                            topMargin=2.2*cm,bottomMargin=2.2*cm)
    S = pdf_styles(); story = []

    story.append(type_badge_table(S,L("align_badge", lang),TYPE_COLORS.get(doc_type,NAVY)))
    story.append(Spacer(1,14))
    story.append(Paragraph(L("align_title", lang),S["Title"]))
    story.append(Paragraph("ThesisSifu Pro v4.2 — " + L("align_sub", lang),S["Sub"]))
    story.append(HR(c=NAVY,t=1.5,b=4,a=10))
    story.append(meta_box(S,[("Document",filename),("Audit ID",audit_id),
        ("Date",datetime.now().strftime("%d %B %Y, %H:%M")),
        ("Report","4 of 4 — Structural Alignment Matrix")]))
    story.append(Spacer(1,10))

    gt = align.golden_thread_score
    gt_color = {"STRONG":GREEN,"ACCEPTABLE":AMBER,"WEAK":RED,"BROKEN":RED}.get(gt,ACCENT)
    story.append(vbanner(S,f"{L('golden_thread', lang)}: {gt}",gt_color,
        subtitle="Problem -> Gap -> RQ -> Theory -> Method -> Analysis -> Finding -> Discussion -> Conclusion"))
    story.append(Spacer(1,10))

    story.append(Paragraph(L("overall_verdict", lang),S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))
    story.append(_safe_para(align.overall_verdict or "(no verdict generated)",S["Body"]))
    story.append(Spacer(1,8))

    story.append(Paragraph(L("spine_ref", lang),S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))
    story.append(spine_box(S,spine,lang)); story.append(Spacer(1,10))

    story.append(PageBreak())
    story.append(Paragraph(L("rq_trace", lang),S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))

    if not align.rows:
        story.append(_safe_para("<i>No alignment rows produced. "
            "This usually means the spine extraction could not identify Research Questions. "
            "Check that the document has clear RQ/RO statements and the GEMINI_API_KEY is valid.</i>",S["Body"]))
    else:
        headers = ["RQ","RO / H","Method + Analysis","Finding","Conclusion","Status"]
        td = [[Paragraph(f"<b>{h}</b>",S["TblH"]) for h in headers]]
        for r in align.rows:
            sc = ALIGN_COLORS.get(r.status,ACCENT)
            roh = r.ro + (f" / {r.hypothesis}" if r.hypothesis and r.hypothesis!="—" else "")
            td.append([
                _safe_para(r.rq[:100],S["TblCSm"]),
                _safe_para(roh[:90],S["TblCSm"]),
                _safe_para(f"{r.method[:60]}\n{r.analysis[:60]}",S["TblCSm"]),
                _safe_para(r.finding[:90],S["TblCSm"]),
                _safe_para(r.conclusion[:90],S["TblCSm"]),
                Paragraph(f'<font color="{sc.hexval()}"><b>{r.status}</b></font>',S["TblCSm"]),
            ])
        cw = W-4*cm
        t = Table(td,colWidths=[cw*.18,cw*.15,cw*.20,cw*.17,cw*.17,cw*.13], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),
            ("BOX",(0,0),(-1,-1),.5,RULE),("INNERGRID",(0,0),(-1,-1),.3,RULE),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
            ("VALIGN",(0,0),(-1,-1),"TOP")]))
        story.append(t); story.append(Spacer(1,10))

        story.append(Paragraph(L("row_notes", lang),S["SecHead"]))
        story.append(HR(c=NAVY,t=1,b=0,a=6))
        for i,r in enumerate(align.rows,1):
            sc = ALIGN_COLORS.get(r.status,ACCENT)
            badge = Table([[Paragraph(r.status,S["Badge"])]],colWidths=[3*cm])
            badge.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),sc),
                ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3)]))
            hrow = Table([[_safe_para(f"<b>RQ{i}.</b> {r.rq[:140]}",S["Body"]),badge]],
                          colWidths=[W-5*cm-3.2*cm,3*cm])
            hrow.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),
                ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))
            story.append(hrow)
            trace = [("Theory / construct", r.theory_construct), ("Instrument", r.instrument),
                     ("Discussion", r.discussion)]
            trace_txt = " | ".join(f"<b>{k}:</b> {v}" for k, v in trace if v and v != "—")
            if trace_txt: story.append(_safe_para(trace_txt, S["TblCSm"]))
            if r.broken_link: story.append(_safe_para(f"<b>{L('broken_link', lang)}:</b> {r.broken_link}", S["Detail"]))
            if r.evidence:    story.append(_safe_para(f"<b>{L('evidence', lang)}:</b> <i>{r.evidence}</i>", S["Excerpt"]))
            if r.note:        story.append(_safe_para(f"<i>{r.note}</i>",S["Excerpt"]))
            if r.fix:         story.append(_safe_para(f"<b>{L('fix', lang)}:</b> {r.fix}", S["Detail"]))
            story.append(Spacer(1,6))

    def _status_table(headers, rows, widths):
        td = [[Paragraph(f"<b>{h}</b>", S["TblH"]) for h in headers]]
        for cells, status in rows:
            sc = ALIGN_COLORS.get(status, ACCENT)
            td.append([_safe_para(str(x), S["TblCSm"]) for x in cells] +
                      [Paragraph(f'<font color="{sc.hexval()}"><b>{status}</b></font>', S["TblCSm"])])
        cw = W-4*cm
        t = Table(td, colWidths=[cw*w for w in widths], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),
            ("BOX",(0,0),(-1,-1),.5,RULE),("INNERGRID",(0,0),(-1,-1),.3,RULE),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
            ("VALIGN",(0,0),(-1,-1),"TOP")]))
        return t

    if align.chapter_links:
        story.append(PageBreak())
        story.append(Paragraph(L("chapter_links", lang), S["SecHead"]))
        story.append(HR(c=NAVY,t=1,b=0,a=6))
        story.append(_status_table(
            ["Link", L("evidence", lang), L("fix", lang), "Status"],
            [((f"{x.get('from_chapter','')} → {x.get('to_chapter','')}: {x.get('link','')}",
               x.get("evidence",""), x.get("fix","")), str(x.get("status","UNCLEAR")).upper())
             for x in align.chapter_links],
            [.27, .33, .27, .13]))
        story.append(Spacer(1,10))

    if align.construct_consistency:
        story.append(Paragraph(L("constructs", lang), S["SecHead"]))
        story.append(HR(c=NAVY,t=1,b=0,a=6))
        story.append(_status_table(
            ["Construct", "Trace", "Note", "Status"],
            [((x.get("construct",""), x.get("trace",""), x.get("note","")),
              str(x.get("status","UNCLEAR")).upper())
             for x in align.construct_consistency],
            [.18, .42, .27, .13]))
        story.append(Spacer(1,10))

    if align.critical_gaps:
        story.append(PageBreak())
        story.append(Paragraph(L("critical_gaps", lang),S["SecHead"]))
        story.append(HR(c=NAVY,t=1,b=0,a=6))
        for g in align.critical_gaps:
            story.append(_safe_para(f"• {g}",S["Bullet"]))
        story.append(Spacer(1,10))

    if align.structural_recommendations:
        story.append(Paragraph(L("struct_recs", lang),S["SecHead"]))
        story.append(HR(c=NAVY,t=1,b=0,a=6))
        for rec in align.structural_recommendations:
            story.append(_safe_para(f"• {rec}",S["Bullet"]))
        story.append(Spacer(1,10))

    # Chapter delivery scorecard
    story.append(PageBreak())
    story.append(Paragraph(L("scorecard", lang),S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))
    story.append(Paragraph("Did each subsection deliver on its expected purpose? "
        "High critical counts against a subsection = it did not.",S["Body"]))
    story.append(Spacer(1,6))
    td = [[Paragraph(f"<b>{h}</b>",S["TblH"]) for h in
           ["Chapter / Subsection","Expected Purpose","Crit","Mod","Sugg"]]]
    for cs in chs:
        nc2=sum(1 for c in cs.comments if c.severity=="CRITICAL")
        nm2=sum(1 for c in cs.comments if c.severity=="MODERATE")
        ns2=sum(1 for c in cs.comments if c.severity=="SUGGESTION")
        td.append([_safe_para(f"<b>Ch.{cs.chapter_num} — {_shorten(cs.chapter_title, 50)}</b>",S["TblC"]),
                   Paragraph("(see subsections)",S["TblCSm"]),
                   Paragraph(f'<font color="{RED.hexval()}">{nc2}</font>' if nc2 else "0",S["TblC"]),
                   Paragraph(f'<font color="{AMBER.hexval()}">{nm2}</font>' if nm2 else "0",S["TblC"]),
                   Paragraph(str(ns2),S["TblC"])])
        for sub in cs.subsections:
            snc=sum(1 for c in sub.comments if c.severity=="CRITICAL")
            snm=sum(1 for c in sub.comments if c.severity=="MODERATE")
            sns=len(sub.comments)-snc-snm
            td.append([_safe_para(f"   §{sub.subsection_num} {_shorten(sub.title, 50)}",S["TblCSm"]),
                       _safe_para(sub.expected_purpose[:120],S["TblCSm"]),
                       Paragraph(f'<font color="{RED.hexval()}">{snc}</font>' if snc else "0",S["TblCSm"]),
                       Paragraph(f'<font color="{AMBER.hexval()}">{snm}</font>' if snm else "0",S["TblCSm"]),
                       Paragraph(str(sns),S["TblCSm"])])
    cw = W-4*cm
    t = Table(td, colWidths=[cw*.32, cw*.42, cw*.08, cw*.08, cw*.08], repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),("BOX",(0,0),(-1,-1),.5,RULE),
        ("INNERGRID",(0,0),(-1,-1),.3,RULE),("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(t); story.append(Spacer(1,16)); story.append(HR())
    story.append(Paragraph("ThesisSifu Pro v4.2 — Alignment Matrix. "
        "Use alongside Report 1, 2, 3 to triangulate revision priorities.",S["Footer"]))
    doc.build(story); buf.seek(0); return buf.getvalue()


# ── Global audit lock ──────────────────────────────────────────
# One audit at a time — prevents two concurrent pdfplumber+fitz+Gemini
# pipelines from doubling memory and causing OOM kill on Railway.
# Lazy init: asyncio.Semaphore MUST be created inside a running event loop
# (Python 3.10+ deprecation, Python 3.13 enforcement).
_AUDIT_LOCK: Optional[asyncio.Semaphore] = None

def _get_audit_lock() -> asyncio.Semaphore:
    global _AUDIT_LOCK
    if _AUDIT_LOCK is None:
        _AUDIT_LOCK = asyncio.Semaphore(2)  # Level 2: 2 concurrent audits (requires 2GB RAM on Railway)
    return _AUDIT_LOCK

# ── Audit pipeline (shared by /audit and /audit/stream) ────────
def _progress(stage: str, message: str, pct: int):
    return ("progress", {"stage": stage, "message": message, "pct": pct})

async def _audit_pipeline(content: bytes, filename: str, audit_id: str):
    """Async generator: yields ("progress", {...}) events, then ("result", {...})."""
    import gc

    yield _progress("extract", "Extracting text from document...", 5)
    full_text = extract_text(content, filename)
    print(f"[ThesisSifu] Extracted {len(full_text)} chars from {filename}")
    if not full_text.strip():
        full_text = "[Document appears image-based or empty — text extraction returned nothing]"
    lang_info = detect_language(full_text)
    lang = lang_info["lang"]
    lang_name = "Bahasa Melayu" if lang == "ms" else "English"
    print(f"[ThesisSifu] Language: {lang} (BM ratio {lang_info['bm_ratio']}, abstracts excluded)")
    yield _progress("extract", f"Extracted {len(full_text):,} characters — thesis language: {lang_name}", 10)

    chs = split_chapters(full_text)
    for cs in chs:
        cs.subsections = split_subsections(cs)
    total_subs = sum(len(cs.subsections) for cs in chs)
    print(f"[ThesisSifu] Split: {len(chs)} chapters, {total_subs} subsections")
    yield _progress("split", f"Found {len(chs)} chapters, {total_subs} subsections", 14)

    yield _progress("spine", "Extracting thesis spine and mapping what each chapter commits to...", 16)
    spine, clf, _ = await asyncio.gather(
        extract_spine(full_text, chs), classify_document(full_text), digest_chapters(chs))
    doc_type = clf.get("type", "MASTERS")
    field    = clf.get("field", "UNKNOWN")
    inst     = clf.get("institution", "UNKNOWN")
    if spine.title == "UNKNOWN" and clf.get("title", "UNKNOWN") != "UNKNOWN":
        spine.title = clf["title"]
    digested = sum(1 for cs in chs if cs.digest)
    print(f"[ThesisSifu] doc_type={doc_type} title={spine.title[:60]} digests={digested}/{len(chs)}")
    yield _progress("spine", f"Spine extracted — {len(spine.research_questions)} RQs found; "
                             f"golden-thread map built for {digested}/{len(chs)} chapters", 26)

    thread = thread_context(chs)
    sem = asyncio.Semaphore(10)
    async def bounded(sub, cs):
        async with sem:
            return sub, await audit_subsection(sub, cs, doc_type, spine, full_text, thread, lang)
    tasks = [bounded(sub, cs) for cs in chs for sub in cs.subsections]
    yield _progress("audit", f"Auditing {len(tasks)} subsections in parallel...", 28)
    done, step = 0, max(1, len(tasks) // 8)
    for fut in asyncio.as_completed(tasks):
        sub, cmts = await fut
        sub.comments = cmts
        done += 1
        if done % step == 0 or done == len(tasks):
            yield _progress("audit", f"Audited {done}/{len(tasks)} subsections...",
                            28 + int(done / len(tasks) * 34))
    for cs in chs:
        cs.comments = [c for sub in cs.subsections for c in sub.comments]
    del tasks
    gc.collect()
    total_comments = sum(len(cs.comments) for cs in chs)
    print(f"[ThesisSifu] Total paragraph comments: {total_comments}")
    yield _progress("audit", f"Paragraph audit complete — {total_comments} comments", 62)

    yield _progress("analysis", "Tracing the golden thread across chapters and retrieving "
                                "supporting literature...", 64)
    align, verified_refs = await asyncio.gather(
        run_alignment_audit(spine, chs, lang), attach_literature(chs, lang))
    print(f"[ThesisSifu] Alignment score={align.golden_thread_score} refs={len(verified_refs)}")
    yield _progress("analysis", f"Golden thread: {align.golden_thread_score}; "
                                f"{len(verified_refs)} verified references found", 76)

    yield _progress("analysis", "Writing the examiner report...", 78)
    examiner_text = await run_examiner(doc_type, spine, field, inst, chs, align, verified_refs, lang)
    gc.collect()

    yield _progress("pdf", "Building PDF reports...", 88)
    pdf1 = build_examiner_pdf(filename, audit_id, doc_type, clf, spine, examiner_text, align, chs, lang)
    pdf3 = build_commentary_pdf(filename, audit_id, doc_type, clf, spine, chs, lang)
    pdf4 = build_alignment_pdf(filename, audit_id, doc_type, clf, spine, align, chs, lang)
    if filename.lower().endswith(".pdf"):
        output2   = build_annotated_pdf(content, full_text, audit_id, chs, lang)
        out2_name = "2_Annotated_Thesis.pdf"
    else:
        output2   = build_annotated_docx(content, filename, audit_id, chs, clf, lang)
        out2_name = "2_Annotated_Thesis.docx"
    del content, chs, examiner_text, align
    gc.collect()

    yield _progress("zip", "Packaging ZIP...", 95)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("1_Examiner_Audit_Report.pdf", pdf1)
        zf.writestr(out2_name, output2)
        zf.writestr("3_Commentary_Report.pdf", pdf3)
        zf.writestr("4_Alignment_Matrix_Report.pdf", pdf4)
    print(f"[ThesisSifu] Audit {audit_id} complete. ZIP ready.")
    yield ("result", {"zip_bytes": buf.getvalue(), "comments": total_comments, "lang": lang})


# ── Main endpoint ──────────────────────────────────────────────
@app.post("/audit")
async def audit_document(file: UploadFile = File(...)):
    content  = await file.read()
    filename = file.filename or "document"

    if not filename.lower().endswith((".pdf",".docx")):
        raise HTTPException(400, "Only PDF and DOCX files are supported.")

    audit_id = "SUP-" + hashlib.md5(content).hexdigest()[:10].upper()
    print(f"[ThesisSifu] Starting audit {audit_id} for {filename}")

    async with _get_audit_lock():
        result = None
        async for kind, data in _audit_pipeline(content, filename, audit_id):
            if kind == "result":
                result = data

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    tmp.write(result["zip_bytes"]); tmp.close()
    return FileResponse(
        path=tmp.name, media_type="application/zip",
        filename=f"ThesisSifu_v4_Audit_{audit_id}.zip",
        background=BackgroundTask(lambda: os.unlink(tmp.name)),
        headers={"Access-Control-Expose-Headers":"Content-Disposition"},
    )


# ── In-memory result store ─────────────────────────────────────
# Stores completed ZIP bytes keyed by audit_id for ~10 minutes
# so the browser can download after SSE stream ends.
_RESULTS: dict[str, bytes] = {}
_RESULT_LOCK = None  # lazy init like _AUDIT_LOCK

def _get_result_lock():
    global _RESULT_LOCK
    if _RESULT_LOCK is None:
        _RESULT_LOCK = asyncio.Lock()
    return _RESULT_LOCK


def _sse(event: str, data: dict) -> str:
    """Format a Server-Sent Event string."""
    return f"event: {event}\ndata: {json.dumps(data)}\n\n"

_KEEPALIVE_SECS = 15

async def _await_with_keepalive(aw):
    """Await `aw`, yielding SSE comment lines every few seconds meanwhile.
    Long stages (chapter digests, alignment, examiner) can run for minutes without
    a progress event; proxies drop idle connections, so we keep bytes flowing.
    The awaited result is delivered as the final item: ("value", result)."""
    fut = asyncio.ensure_future(aw)
    try:
        while True:
            done, _ = await asyncio.wait({fut}, timeout=_KEEPALIVE_SECS)
            if done:
                yield ("value", fut.result())
                return
            yield ("keepalive", ": keepalive\n\n")
    finally:
        if not fut.done():      # client disconnected: stop the work
            fut.cancel()


@app.post("/audit/stream")
async def audit_stream(file: UploadFile = File(...)):
    """SSE streaming audit endpoint.
    Returns a text/event-stream with progress events, then a 'done' event
    containing the audit_id. The client then fetches /audit/download/{audit_id}
    to get the ZIP.

    Events emitted:
      progress  { stage: str, message: str, pct: int }
      error     { message: str }
      done      { audit_id: str, filename: str, comments: int }
    Comment lines (": keepalive") are sent during long stages.
    """
    content  = await file.read()
    filename = file.filename or "document"

    if not filename.lower().endswith((".pdf", ".docx")):
        async def _err():
            yield _sse("error", {"message": "Only PDF and DOCX files are supported."})
        return StreamingResponse(_err(), media_type="text/event-stream")

    audit_id = "SUP-" + hashlib.md5(content).hexdigest()[:10].upper()

    async def _stream() -> AsyncGenerator[str, None]:
        lock = _get_audit_lock()
        acquired = False
        agen = _audit_pipeline(content, filename, audit_id)
        try:
            if lock.locked():
                yield _sse("progress", {"stage": "queue", "message": "Waiting for an audit slot...", "pct": 1})
            async for kind, val in _await_with_keepalive(lock.acquire()):
                if kind == "keepalive":
                    yield val
            acquired = True

            while True:
                item = None
                async for kind, val in _await_with_keepalive(agen.__anext__()):
                    if kind == "keepalive":
                        yield val
                    else:
                        item = val
                event, data = item
                if event == "progress":
                    yield _sse("progress", data)
                    continue
                # event == "result"
                async with _get_result_lock():
                    _RESULTS[audit_id] = data["zip_bytes"]

                async def _cleanup():
                    await asyncio.sleep(600)
                    async with _get_result_lock():
                        _RESULTS.pop(audit_id, None)
                asyncio.create_task(_cleanup())

                yield _sse("progress", {"stage": "done", "message": "Audit complete!", "pct": 100})
                yield _sse("done", {
                    "audit_id": audit_id,
                    "filename": f"ThesisSifu_v4_Audit_{audit_id}.zip",
                    "comments": data["comments"],
                })
                break
        except Exception as e:
            print(f"[ThesisSifu] Stream error: {e}\n{traceback.format_exc()}")
            yield _sse("error", {"message": f"Audit failed: {str(e)[:200]}"})
        finally:
            try:
                await agen.aclose()
            except Exception:
                pass
            if acquired:
                lock.release()

    return StreamingResponse(
        _stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",       # disable Nginx/Railway buffering
            "Access-Control-Allow-Origin": "*",
        }
    )


@app.get("/audit/download/{audit_id}")
async def audit_download(audit_id: str):
    """Download the ZIP for a completed audit by audit_id."""
    async with _get_result_lock():
        zip_bytes = _RESULTS.get(audit_id)
    if not zip_bytes:
        raise HTTPException(404, "Audit result not found or expired (results kept for 10 minutes).")
    return StreamingResponse(
        io.BytesIO(zip_bytes),
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="ThesisSifu_v4_Audit_{audit_id}.zip"',
            "Access-Control-Expose-Headers": "Content-Disposition",
        }
    )


@app.get("/health")
async def health():
    return {
        "status":                "ok",
        "gemini_available":      gemini_client is not None,
        "gemini_model":          GEMINI_MODEL,
        "claude_available":      claude_client is not None,
        "claude_model":          CLAUDE_MODEL,
        "literature_source":     "openalex" if OPENALEX_API_KEY else "crossref",
        "version":               "4.2.0",
        "outputs": ["1_Examiner_Audit_Report.pdf","2_Annotated_Thesis.(docx|pdf)",
                    "3_Commentary_Report.pdf","4_Alignment_Matrix_Report.pdf"],
    }
