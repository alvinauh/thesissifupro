"""
ThesisSifu Pro v4.0 — Multi-Agent Thesis Panel (Alignment-Aware Edition)
========================================================================
Four-output audit system with structural-spine extraction, subsection-level
analysis, dedicated alignment auditing, and a curated framework library.

Pipeline:
  Stage 0  Spine Extraction        (Gemini)  → ThesisSpine
  Stage 1  Chapter + Subsection Split        → structural map
  Stage 2  Subsection paragraph audit (Gemini, parallel) → uses spine + framework library
  Stage 3  Alignment Audit         (Claude)  → AlignmentMatrix
  Stage 4  Holistic Examiner Report (Claude) → uses spine + matrix

Outputs (ZIP):
  1_Examiner_Audit_Report.pdf     — holistic critique, references spine
  2_Annotated_Thesis.(docx|pdf)   — subsection-aware inline comments
  3_Commentary_Report.pdf         — grouped by chapter → subsection → paragraph
  4_Alignment_Matrix_Report.pdf   — PS↔RQ↔RO↔Method↔Analysis↔Findings↔Conclusion

Endpoint:  POST /audit   multipart/form-data { file: <pdf|docx> }
Returns:   application/zip
"""

from __future__ import annotations

import io, os, re, json, asyncio, zipfile, hashlib, tempfile
from datetime import datetime
from dataclasses import dataclass, field as dc_field
from typing import Optional

import pypdf
import fitz  # PyMuPDF
from docx import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.opc.part import XmlPart
from docx.opc.packuri import PackURI

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
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
app = FastAPI(title="ThesisSifu Pro v4 — Alignment-Aware", version="4.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"],
                   allow_credentials=False, allow_methods=["*"], allow_headers=["*"])


# ── AI Clients ─────────────────────────────────────────────────
claude_client     = None

if GEMINI_AVAILABLE:
    gkey = os.environ.get("GEMINI_API_KEY")
    if gkey:
        gemini_client = genai_sdk.Client(api_key=gkey)
        GEMINI_MODEL  = "gemini-2.0-flash-lite"
        print(f"Gemini {GEMINI_MODEL} ready (google.genai SDK)")
    else:
        gemini_client = None
        GEMINI_MODEL  = ""
else:
    gemini_client = None
    GEMINI_MODEL  = ""

if ANTHROPIC_AVAILABLE:
    akey = os.environ.get("ANTHROPIC_API_KEY")
    if akey:
        claude_client = anthropic.AsyncAnthropic(api_key=akey)
        print("Claude Sonnet 4.6 ready")


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

# Alignment status colours
ALIGN_COLORS = {
    "ALIGNED":     GREEN,
    "PARTIAL":     AMBER,
    "MISALIGNED":  RED,
    "MISSING":     colors.HexColor("#555555"),
    "UNCLEAR":     ACCENT,
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


# ── Canonical Framework & Method Library ──────────────────────
# Curated, named-author references. The LLM is instructed to recommend
# FROM this list rather than invent titles — prevents hallucination while
# still giving students actionable, specific guidance.
CANONICAL_FRAMEWORKS = """
THEORETICAL FRAMEWORKS (recommend by exact name + seminal author):

Technology / Information Systems:
  - Technology Acceptance Model (TAM) — Davis (1989)
  - UTAUT / UTAUT2 — Venkatesh et al. (2003, 2012)
  - Diffusion of Innovations — Rogers (2003)
  - Task-Technology Fit — Goodhue & Thompson (1995)
  - DeLone & McLean IS Success Model — DeLone & McLean (2003)

Behavioural / Psychological:
  - Theory of Planned Behaviour — Ajzen (1991)
  - Theory of Reasoned Action — Fishbein & Ajzen (1975)
  - Social Cognitive Theory — Bandura (1986)
  - Self-Determination Theory — Deci & Ryan (1985, 2000)
  - Health Belief Model — Rosenstock (1974)

Strategic Management / Business:
  - Resource-Based View — Barney (1991)
  - Dynamic Capabilities — Teece, Pisano & Shuen (1997)
  - Porter's Five Forces — Porter (1980)
  - Stakeholder Theory — Freeman (1984)
  - Institutional Theory — DiMaggio & Powell (1983)

Organisational Behaviour / HR:
  - Job Demands-Resources Model — Bakker & Demerouti (2007)
  - Social Exchange Theory — Blau (1964)
  - Transformational Leadership — Bass (1985)
  - Organisational Citizenship Behaviour — Organ (1988)
  - Psychological Capital — Luthans et al. (2007)

Marketing / Consumer Behaviour:
  - Stimulus-Organism-Response (SOR) — Mehrabian & Russell (1974)
  - SERVQUAL — Parasuraman, Zeithaml & Berry (1988)
  - Customer-Based Brand Equity — Keller (1993)
  - Expectation-Confirmation Theory — Oliver (1980)

Education / Learning:
  - Constructivism — Piaget, Vygotsky
  - Bloom's Taxonomy (revised) — Anderson & Krathwohl (2001)
  - Community of Inquiry — Garrison, Anderson & Archer (2000)
  - Self-Regulated Learning — Zimmerman (2002)
  - TPACK — Mishra & Koehler (2006)

Sociology / Public Policy:
  - Structuration Theory — Giddens (1984)
  - Social Capital — Putnam (2000), Bourdieu (1986)
  - Capability Approach — Sen (1999), Nussbaum (2011)
"""

CANONICAL_METHODS = """
METHODOLOGICAL REFERENCES (recommend by exact name + seminal author):

Quantitative analytical techniques:
  - PLS-SEM — Hair, Hult, Ringle & Sarstedt (2017, 2022); for prediction & theory development
  - CB-SEM — Byrne (2016), Kline (2015); for confirmatory theory testing
  - Multiple Regression — Field (2018), Hair et al. (2019)
  - Hierarchical regression / moderation — Aiken & West (1991), Hayes PROCESS (2018)
  - Mediation — Baron & Kenny (1986), Preacher & Hayes (2008), Hayes (2018)
  - ANOVA/MANOVA — Tabachnick & Fidell (2019)
  - Factor Analysis (EFA/CFA) — Hair et al. (2019), Brown (2015)
  - Logistic regression — Hosmer, Lemeshow & Sturdivant (2013)

Qualitative analytical techniques:
  - Thematic Analysis — Braun & Clarke (2006, 2019)
  - Reflexive Thematic Analysis — Braun & Clarke (2022)
  - Grounded Theory — Charmaz (2014), Strauss & Corbin (1998)
  - Gioia Methodology — Gioia, Corley & Hamilton (2013)
  - Interpretative Phenomenological Analysis (IPA) — Smith, Flowers & Larkin (2009)
  - Case Study — Yin (2018), Eisenhardt (1989)
  - Narrative Inquiry — Clandinin & Connelly (2000)
  - Discourse Analysis — Fairclough (2013)

Mixed methods:
  - Mixed Methods Typology — Creswell & Plano Clark (2018)
  - Sequential Explanatory / Exploratory — Creswell (2014)

Systematic / scoping review:
  - PRISMA 2020 — Page et al. (2021)
  - Scoping review — Arksey & O'Malley (2005), Tricco et al. (2018)
  - Bibliometric analysis — Donthu et al. (2021)

Sampling / sample size:
  - Cochran's formula — Cochran (1977)
  - Krejcie & Morgan table — Krejcie & Morgan (1970)
  - Yamane's formula — Yamane (1967)
  - G*Power power analysis — Faul, Erdfelder, Lang & Buchner (2009)
  - Minimum sample for PLS-SEM (10-times rule, inverse square root) — Hair et al. (2017)

Validity & reliability:
  - Cronbach's alpha — Cronbach (1951)
  - Composite reliability, AVE — Hair et al. (2017), Fornell & Larcker (1981)
  - HTMT discriminant validity — Henseler, Ringle & Sarstedt (2015)
  - Common Method Bias — Podsakoff et al. (2003, 2012); Harman's single-factor test
  - Content validity index (CVI) — Lynn (1986), Polit & Beck (2006)

Qualitative trustworthiness:
  - Lincoln & Guba (1985) — credibility, transferability, dependability, confirmability
  - Member checking, audit trail, thick description
"""


# ── Data classes ───────────────────────────────────────────────
@dataclass
class ThesisSpine:
    """The structural backbone of the thesis — extracted in Stage 0
    and used as the reference object for every downstream audit."""
    title:              str = "UNKNOWN"
    problem_statement:  str = "NOT FOUND"
    research_gap:       str = "NOT FOUND"
    research_questions: list = dc_field(default_factory=list)  # list[str]
    research_objectives:list = dc_field(default_factory=list)
    hypotheses:         list = dc_field(default_factory=list)
    theory_used:        str = "NOT FOUND"
    variables:          list = dc_field(default_factory=list)  # constructs/IV/DV/MV
    methodology:        str = "NOT FOUND"  # design + paradigm
    sampling:           str = "NOT FOUND"
    instrument:         str = "NOT FOUND"
    analysis_technique: str = "NOT FOUND"
    key_findings:       list = dc_field(default_factory=list)
    conclusions:        list = dc_field(default_factory=list)
    discipline:         str = "UNKNOWN"

@dataclass
class ParagraphComment:
    chapter:           str
    subsection:        str          # e.g. "3.2" or "—" if no subsection detected
    subsection_title:  str
    para_index:        int
    page_estimate:     int
    para_excerpt:      str
    severity:          str
    issue:             str
    recommendation:    str
    literature_needed: str
    theory_needed:     str
    suggested_framework: str = ""   # NEW: specific named framework from library
    suggested_method:   str = ""    # NEW: specific named method/reference

@dataclass
class Subsection:
    chapter_num:    str
    subsection_num: str
    title:          str
    text:           str
    expected_purpose: str = ""      # what this kind of subsection SHOULD contain
    comments:       list = dc_field(default_factory=list)

@dataclass
class ChapterSummary:
    chapter_num:   str
    chapter_title: str
    text:          str = ""
    subsections:   list = dc_field(default_factory=list)  # list[Subsection]
    comments:      list = dc_field(default_factory=list)  # flat list across subsections

@dataclass
class AlignmentRow:
    """One row in the alignment matrix — e.g. RQ1 → RO1 → Method → Finding → Conclusion."""
    rq:           str = "—"
    ro:           str = "—"
    hypothesis:   str = "—"
    method:       str = "—"
    analysis:     str = "—"
    finding:      str = "—"
    conclusion:   str = "—"
    status:       str = "UNCLEAR"   # ALIGNED|PARTIAL|MISALIGNED|MISSING|UNCLEAR
    note:         str = ""

@dataclass
class AlignmentMatrix:
    rows:               list = dc_field(default_factory=list)  # list[AlignmentRow]
    overall_verdict:    str = ""
    golden_thread_score: str = ""   # STRONG|ACCEPTABLE|WEAK|BROKEN
    critical_gaps:      list = dc_field(default_factory=list)  # list[str]
    structural_recommendations: list = dc_field(default_factory=list)


# ── Text extraction ─────────────────────────────────────────────
def extract_text(content: bytes, filename: str) -> str:
    fn = filename.lower()
    if fn.endswith(".pdf"):
        reader = pypdf.PdfReader(io.BytesIO(content))
        return "\n".join(p.extract_text() or "" for p in reader.pages)
    if fn.endswith(".docx"):
        doc = DocxDocument(io.BytesIO(content))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    return ""


# ── Chapter + Subsection splitter ─────────────────────────────
_CH_PAT = re.compile(
    r'(?:^|\n)(?:CHAPTER\s+(\d+|[IVX]+)|(\d+)\.0)\b[:\s\-—–]*([^\n]{0,80})',
    re.IGNORECASE,
)

# Subsection: matches 1.1, 1.2.3, 3.4 etc. at start of line.
_SUB_PAT = re.compile(
    r'(?:^|\n)[ \t]*(\d+\.\d+(?:\.\d+)?)\b[:\s\-—–]*([^\n]{0,120})',
)

# Expected purpose of common subsection patterns. Keys are heuristic title fragments.
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
    "reliability":          "Report reliability evidence (Cronbach's α, composite reliability).",
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

def _purpose_for(title: str) -> str:
    t = title.lower()
    for k, v in _SUBSECTION_PURPOSE.items():
        if k in t:
            return v
    return "Auditor should infer purpose from chapter context."

def split_chapters(text: str) -> list[ChapterSummary]:
    matches = list(_CH_PAT.finditer(text))
    if not matches:
        # No chapter markers — fall back to fixed-size chunks
        size = 6000
        return [ChapterSummary(chapter_num=str(i+1),
                               chapter_title=f"Section {i+1}",
                               text=text[s:s+size])
                for i, s in enumerate(range(0, len(text), size))]
    chapters = []
    for i, m in enumerate(matches):
        num   = (m.group(1) or m.group(2) or str(i+1)).strip()
        title = (m.group(3) or "").strip() or f"Chapter {num}"
        start = m.start()
        end   = matches[i+1].start() if i+1 < len(matches) else len(text)
        ch_text = text[start:end].strip()
        chapters.append(ChapterSummary(chapter_num=num, chapter_title=title, text=ch_text))
    return chapters or [ChapterSummary("1", "Full Document", text)]

def split_subsections(chapter: ChapterSummary) -> list[Subsection]:
    """Find 1.1, 1.2... within a chapter. If none, treat the whole chapter as one."""
    matches = list(_SUB_PAT.finditer(chapter.text))
    # Filter spurious matches that don't start with the chapter number
    matches = [m for m in matches if m.group(1).split(".")[0] == chapter.chapter_num]
    if not matches:
        return [Subsection(chapter_num=chapter.chapter_num, subsection_num="—",
                           title=chapter.chapter_title, text=chapter.text,
                           expected_purpose=_purpose_for(chapter.chapter_title))]
    subs = []
    for i, m in enumerate(matches):
        sub_num = m.group(1).strip()
        sub_title = (m.group(2) or "").strip() or f"Section {sub_num}"
        start = m.start()
        end   = matches[i+1].start() if i+1 < len(matches) else len(chapter.text)
        sub_text = chapter.text[start:end].strip()
        subs.append(Subsection(
            chapter_num=chapter.chapter_num,
            subsection_num=sub_num,
            title=sub_title,
            text=sub_text,
            expected_purpose=_purpose_for(sub_title),
        ))
    return subs


# ── Stage 0: Spine Extraction (Gemini) ────────────────────────
_SPINE_PROMPT = """
You are extracting the STRUCTURAL SPINE of an academic thesis or article.
Return STRICT JSON only — no markdown, no commentary.

Fields (use "NOT FOUND" or [] if absent from the excerpt):
{{
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
}}

Document excerpt (front + back samples):
\"\"\"{text}\"\"\"
"""

async def extract_spine(text: str) -> ThesisSpine:
    if not gemini_client or not text.strip():
        return ThesisSpine()
    # Sample front (intro/method/RQs) AND back (findings/conclusion) for fuller spine
    front = text[:6000]
    back  = text[-4000:] if len(text) > 10000 else ""
    sample = front + "\n\n[...later in the document...]\n\n" + back
    try:
        r = await gemini_client.aio.models.generate_content(
            model=GEMINI_MODEL,
            contents=_SPINE_PROMPT.format(text=sample[:12000]))
        raw = re.sub(r"```(?:json)?", "", r.text.strip()).strip("`").strip()
        d = json.loads(raw)
        return ThesisSpine(
            title=d.get("title","UNKNOWN"),
            discipline=d.get("discipline","UNKNOWN"),
            problem_statement=d.get("problem_statement","NOT FOUND"),
            research_gap=d.get("research_gap","NOT FOUND"),
            research_questions=d.get("research_questions",[]) or [],
            research_objectives=d.get("research_objectives",[]) or [],
            hypotheses=d.get("hypotheses",[]) or [],
            theory_used=d.get("theory_used","NOT FOUND"),
            variables=d.get("variables",[]) or [],
            methodology=d.get("methodology","NOT FOUND"),
            sampling=d.get("sampling","NOT FOUND"),
            instrument=d.get("instrument","NOT FOUND"),
            analysis_technique=d.get("analysis_technique","NOT FOUND"),
            key_findings=d.get("key_findings",[]) or [],
            conclusions=d.get("conclusions",[]) or [],
        )
    except Exception as e:
        print(f"Spine extraction error: {e}")
        return ThesisSpine()


# ── Stage 1: Classification (unchanged interface; spine carries title) ─
_CLS_PROMPT = """
Classify this academic document. Choose ONE:
JOURNAL_ARTICLE | UNDERGRADUATE | MASTERS | PHD

Return ONLY valid JSON (no markdown):
{{"type":"PHD","confidence":"HIGH","signals":["s1","s2"],"title":"t","authors":"a","field":"f","institution":"i"}}

Excerpt:
\"\"\"{text}\"\"\"
"""

async def classify_document(text: str) -> dict:
    default = {"type":"MASTERS","confidence":"LOW","signals":["fallback"],
               "title":"UNKNOWN","authors":"UNKNOWN","field":"UNKNOWN","institution":"UNKNOWN"}
    if not gemini_client or not text.strip():
        return default
    try:
        r = await gemini_client.aio.models.generate_content(
            model=GEMINI_MODEL,
            contents=_CLS_PROMPT.format(text=text[:5000]))
        raw = re.sub(r"```(?:json)?","", r.text.strip()).strip("`").strip()
        return json.loads(raw)
    except Exception as e:
        print(f"Classification error: {e}")
        return default


# ── Stage 2: Subsection paragraph audit (spine-aware) ─────────
_SUBSECTION_PROMPT = """
You are a PhD supervisor auditing one SUBSECTION of a {doc_type}.

THESIS SPINE (for alignment checking — refer to this in every comment):
  Title:            {title}
  Problem:          {problem}
  Research Qs:      {rqs}
  Research Objs:    {ros}
  Theory used:      {theory}
  Methodology:      {method}
  Analysis:         {analysis}
  Variables:        {vars}

CHAPTER {ch_num} — {ch_title}
SUBSECTION {sub_num} — {sub_title}
Expected purpose of this subsection: {purpose}

YOUR TASK
For 4-8 paragraphs in this subsection, return JSON objects with:
{{
  "para_excerpt": "first 100 chars verbatim",
  "severity": "CRITICAL | MODERATE | SUGGESTION",
  "issue": "specific problem — be concrete and intellectual, not stylistic",
  "recommendation": "specific fix tied to thesis spine and subsection purpose",
  "literature_needed": "type of evidence/studies needed (do NOT invent titles)",
  "theory_needed": "named framework from canonical library below OR 'none'",
  "suggested_framework": "EXACT name + seminal author from canonical list, e.g., 'TAM (Davis 1989)' — empty string if none applies",
  "suggested_method": "EXACT name + seminal author for any methodological recommendation, e.g., 'PLS-SEM (Hair et al. 2017)' — empty string if none applies"
}}

ALIGNMENT FOCUS — flag the following as CRITICAL when present:
  • Paragraph contradicts the problem statement or strays from the RQs/ROs
  • Claim made without subsection delivering on its expected purpose
  • Methodology choice that cannot answer the stated RQ
  • Finding that does not map to any RQ
  • Conclusion that overreaches what the analysis supports
  • Theory invoked but not operationalised in the model/instrument

{frameworks}

{methods}

Return ONLY a JSON array. No markdown fences. No prose before or after.

Subsection text:
\"\"\"{text}\"\"\"
"""

async def audit_subsection(sub: Subsection, ch_title: str, doc_type: str,
                            spine: ThesisSpine) -> list[ParagraphComment]:
    if not gemini_client or not sub.text.strip():
        return []
    prompt = _SUBSECTION_PROMPT.format(
        doc_type=doc_type,
        title=spine.title,
        problem=spine.problem_statement[:300],
        rqs="; ".join(spine.research_questions[:5]) or "NOT FOUND",
        ros="; ".join(spine.research_objectives[:5]) or "NOT FOUND",
        theory=spine.theory_used,
        method=spine.methodology,
        analysis=spine.analysis_technique,
        vars="; ".join(spine.variables[:8]) or "NOT FOUND",
        ch_num=sub.chapter_num, ch_title=ch_title,
        sub_num=sub.subsection_num, sub_title=sub.title,
        purpose=sub.expected_purpose,
        frameworks=CANONICAL_FRAMEWORKS,
        methods=CANONICAL_METHODS,
        text=sub.text[:7000],
    )
    try:
        r = await gemini_client.aio.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt)
        raw = re.sub(r"```(?:json)?", "", r.text.strip()).strip("`").strip()
        items = json.loads(raw)
        if not isinstance(items, list):
            items = [items]
        comments = []
        for idx, it in enumerate(items):
            excerpt = it.get("para_excerpt","")
            pos = sub.text.find(excerpt[:40])
            pg  = max(1, len(sub.text[:pos].split()) // 350 + 1) if pos > 0 else 1
            comments.append(ParagraphComment(
                chapter=f"Chapter {sub.chapter_num}",
                subsection=sub.subsection_num,
                subsection_title=sub.title,
                para_index=idx, page_estimate=pg,
                para_excerpt=excerpt[:150],
                severity=it.get("severity","MODERATE").upper(),
                issue=it.get("issue",""),
                recommendation=it.get("recommendation",""),
                literature_needed=it.get("literature_needed",""),
                theory_needed=it.get("theory_needed","none"),
                suggested_framework=it.get("suggested_framework",""),
                suggested_method=it.get("suggested_method",""),
            ))
        return comments
    except Exception as e:
        print(f"Subsection audit error {sub.chapter_num}.{sub.subsection_num}: {e}")
        return []


# ── Stage 3: Alignment Audit (Claude) ─────────────────────────
_ALIGNMENT_SYSTEM = """
You are a senior academic examiner auditing the STRUCTURAL ALIGNMENT of a thesis.
Your task is to verify the 'golden thread': Problem → RQ → RO → Theory → Method →
Analysis → Findings → Discussion → Conclusion.

Be rigorous and specific. Quote the spine fields. Flag broken links explicitly.
Return STRICT JSON only. No markdown.
"""

_ALIGNMENT_PROMPT = """
THESIS SPINE
============
Title:                {title}
Discipline:           {discipline}
Problem statement:    {problem}
Research gap:         {gap}
Research questions:   {rqs}
Research objectives:  {ros}
Hypotheses:           {hyps}
Theory used:          {theory}
Variables/constructs: {vars}
Methodology:          {method}
Sampling:             {sampling}
Instrument:           {instrument}
Analysis technique:   {analysis}
Key findings:         {findings}
Conclusions:          {conclusions}

CRITICAL ISSUES RAISED BY SUBSECTION AUDIT
==========================================
{critical_issues}

YOUR TASK
=========
Produce a JSON object with:

{{
  "rows": [
    {{
      "rq": "RQ1 verbatim or paraphrase",
      "ro": "RO1 matched to RQ1",
      "hypothesis": "H1 if applicable, else '—'",
      "method": "specific method used to answer RQ1",
      "analysis": "specific analysis technique applied",
      "finding": "what was found relevant to RQ1",
      "conclusion": "conclusion drawn for RQ1",
      "status": "ALIGNED | PARTIAL | MISALIGNED | MISSING | UNCLEAR",
      "note": "concrete reason for this status — quote spine fields if possible"
    }}
    // one row per research question
  ],
  "golden_thread_score": "STRONG | ACCEPTABLE | WEAK | BROKEN",
  "overall_verdict": "2-3 sentence verdict on structural coherence of the thesis",
  "critical_gaps": [
    "specific gap 1 — e.g., 'RQ2 is not answered by the analysis technique used (regression cannot test the mediation claimed)'",
    "specific gap 2"
  ],
  "structural_recommendations": [
    "specific structural fix 1 — name the framework/method to use",
    "specific structural fix 2"
  ]
}}

Return ONLY the JSON object. No prose, no markdown fences.
"""

async def run_alignment_audit(spine: ThesisSpine, chs: list[ChapterSummary]) -> AlignmentMatrix:
    if not claude_client:
        return AlignmentMatrix(
            overall_verdict="Alignment audit unavailable — ANTHROPIC_API_KEY not set.",
            golden_thread_score="UNCLEAR")
    # Compile critical issues from subsection audit for context
    critical_issues = []
    for cs in chs:
        for c in cs.comments:
            if c.severity == "CRITICAL":
                critical_issues.append(f"Ch.{c.chapter} §{c.subsection}: {c.issue}")
    critical_block = "\n".join(critical_issues[:30]) or "(no critical issues flagged at subsection level)"

    prompt = _ALIGNMENT_PROMPT.format(
        title=spine.title, discipline=spine.discipline,
        problem=spine.problem_statement, gap=spine.research_gap,
        rqs="; ".join(spine.research_questions) or "NOT FOUND",
        ros="; ".join(spine.research_objectives) or "NOT FOUND",
        hyps="; ".join(spine.hypotheses) or "NOT FOUND",
        theory=spine.theory_used,
        vars="; ".join(spine.variables) or "NOT FOUND",
        method=spine.methodology, sampling=spine.sampling,
        instrument=spine.instrument, analysis=spine.analysis_technique,
        findings="; ".join(spine.key_findings) or "NOT FOUND",
        conclusions="; ".join(spine.conclusions) or "NOT FOUND",
        critical_issues=critical_block,
    )
    try:
        msg = await claude_client.messages.create(
            model="claude-sonnet-4-6", max_tokens=4096, temperature=0.2,
            system=_ALIGNMENT_SYSTEM,
            messages=[{"role":"user","content":prompt}])
        raw = msg.content[0].text.strip()
        raw = re.sub(r"```(?:json)?","", raw).strip("`").strip()
        d = json.loads(raw)
        rows = [AlignmentRow(
            rq=r.get("rq","—"), ro=r.get("ro","—"),
            hypothesis=r.get("hypothesis","—"), method=r.get("method","—"),
            analysis=r.get("analysis","—"), finding=r.get("finding","—"),
            conclusion=r.get("conclusion","—"),
            status=r.get("status","UNCLEAR").upper(),
            note=r.get("note","")) for r in d.get("rows",[])]
        return AlignmentMatrix(
            rows=rows,
            overall_verdict=d.get("overall_verdict",""),
            golden_thread_score=d.get("golden_thread_score","UNCLEAR").upper(),
            critical_gaps=d.get("critical_gaps",[]),
            structural_recommendations=d.get("structural_recommendations",[]),
        )
    except Exception as e:
        print(f"Alignment audit error: {e}")
        return AlignmentMatrix(
            overall_verdict=f"Alignment audit failed: {e}",
            golden_thread_score="UNCLEAR")


# ── Stage 4: Holistic Examiner (Claude, now spine-aware) ──────
_EXAMINER_SYSTEM = """
You are a senior academic examiner with 25 years of experience.
Produce rigorous, specific, honest examiner-level critique anchored to the
provided thesis spine and alignment matrix. You are NOT a grammar checker —
you are an intellectual critic. Cite specific spine fields when you make claims.
Recommend named frameworks (with seminal author) where appropriate.
Respond in the same language as the document.
"""

_EXAMINER_PROMPTS = {
"PHD": """
External examiner for a PhD viva.

THESIS SPINE
============
Title:            "{title}"
Field:            {field} | Institution: {institution}
Problem:          {problem}
Research Qs:      {rqs}
Research Objs:    {ros}
Theory:           {theory}
Methodology:      {method}
Analysis:         {analysis}

ALIGNMENT VERDICT (from dedicated audit)
========================================
Golden thread score: {gt_score}
Overall:             {align_verdict}
Critical gaps:       {gaps}

CHAPTER + SUBSECTION ISSUE SUMMARY
===================================
{summaries}

CANONICAL LIBRARIES (recommend FROM these — do NOT invent references)
====================================================================
{frameworks}
{methods}

Write a full examiner's report with EXACTLY these sections:

SECTION 1 — ORIGINAL CONTRIBUTION TO KNOWLEDGE
SECTION 2 — GOLDEN THREAD ANALYSIS
Use the alignment matrix above. Quote spine fields. Identify exact broken links.
SECTION 3 — CHAPTER + SUBSECTION ALIGNMENT AUDIT
For each chapter, list which subsections under-delivered on their expected purpose.
SECTION 4 — THEORETICAL FRAMEWORK COHERENCE
Name a specific framework (with seminal author) the candidate should add or strengthen.
SECTION 5 — METHODOLOGICAL RIGOUR
Recommend specific named methods (e.g., PLS-SEM via Hair et al. 2017) if appropriate.
SECTION 6 — DATA QUALITY & ANALYTICAL DEPTH
SECTION 7 — SCOPUS-LEVEL LANGUAGE & TONE
Assess against Q1/Q2 journal standards. Flag specific passages and rewrite needs.
SECTION 8 — CITATION & REFERENCE INTEGRITY
SECTION 9 — CRITICAL VIVA QUESTIONS
List 8 specific questions tied directly to spine and alignment gaps.
SECTION 10 — EXAMINER'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS REQUIRED / REFER (RESUBMIT) / FAIL
Formal statement 6-8 sentences. Specific conditions before degree is awarded.
""",
"MASTERS": """
Internal reader for a Master's thesis viva.

THESIS SPINE
============
Title:        "{title}"
Field:        {field} | Institution: {institution}
Problem:      {problem}
Research Qs:  {rqs}
Research Objs:{ros}
Theory:       {theory}
Methodology:  {method}
Analysis:     {analysis}

ALIGNMENT VERDICT
=================
Golden thread score: {gt_score}
Overall:             {align_verdict}
Critical gaps:       {gaps}

CHAPTER + SUBSECTION ISSUE SUMMARY
==================================
{summaries}

CANONICAL LIBRARIES (recommend FROM these only)
================================================
{frameworks}
{methods}

Write a report with EXACTLY these sections:

SECTION 1 — THESIS ALIGNMENT & GOLDEN THREAD (use the matrix above)
SECTION 2 — CHAPTER + SUBSECTION ALIGNMENT AUDIT
SECTION 3 — THEORETICAL FRAMEWORK (recommend named theory + seminal author)
SECTION 4 — METHODOLOGICAL RIGOUR (recommend named method + seminal author)
SECTION 5 — DATA ROBUSTNESS & INSTRUMENTATION
SECTION 6 — SCOPUS-LEVEL LANGUAGE & TONE
SECTION 7 — CITATION INTEGRITY
SECTION 8 — CRITICAL VIVA QUESTIONS (5 questions tied to alignment gaps)
SECTION 9 — EXAMINER'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS REQUIRED / RESUBMIT
Top 5 issues that must be resolved.
""",
"JOURNAL_ARTICLE": """
Peer reviewer for a Scopus journal.

ARTICLE SPINE
=============
Title:        "{title}"
Field:        {field}
Problem:      {problem}
Research Qs:  {rqs}
Theory:       {theory}
Methodology:  {method}
Analysis:     {analysis}

ALIGNMENT VERDICT
=================
Golden thread score: {gt_score}
Overall:             {align_verdict}
Critical gaps:       {gaps}

SECTION ISSUE SUMMARY
=====================
{summaries}

CANONICAL LIBRARIES
===================
{frameworks}
{methods}

Write a review with EXACTLY these sections:

SECTION 1 — CONTRIBUTION & NOVELTY
SECTION 2 — LITERATURE CURRENCY & GAPS
SECTION 3 — METHODOLOGY (cite named methodological references)
SECTION 4 — RESULTS & ANALYSIS
SECTION 5 — ARGUMENT COHERENCE (use alignment matrix above)
SECTION 6 — SCOPUS-LEVEL LANGUAGE & TONE
SECTION 7 — CITATION INTEGRITY
SECTION 8 — PUBLICATION VERDICT
One of: ACCEPT / MINOR REVISIONS / MAJOR REVISIONS / REJECT
Top 3 critical revisions.
""",
"UNDERGRADUATE": """
Supervisor reviewing an undergraduate FYP.

PROJECT SPINE
=============
Title:        "{title}"
Field:        {field}
Problem:      {problem}
Research Qs:  {rqs}
Methodology:  {method}
Analysis:     {analysis}

ALIGNMENT VERDICT
=================
Golden thread score: {gt_score}
Overall:             {align_verdict}

SECTION ISSUE SUMMARY
=====================
{summaries}

CANONICAL LIBRARIES (recommend FROM these — appropriate for undergraduate level)
=================================================================================
{frameworks}
{methods}

Write a report with EXACTLY these sections:

SECTION 1 — SCOPE & RESEARCH QUESTION (alignment with project objectives)
SECTION 2 — LITERATURE REVIEW (recommend named theory if missing)
SECTION 3 — METHODOLOGY (recommend named method if appropriate)
SECTION 4 — ANALYSIS & FINDINGS
SECTION 5 — CRITICAL THINKING
SECTION 6 — WRITING & ACADEMIC CONVENTIONS
SECTION 7 — TOP ISSUES TO FIX
SECTION 8 — SUPERVISOR'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS / FAIL
""",
}

async def run_examiner(doc_type, spine: ThesisSpine, field, inst,
                       chs: list[ChapterSummary], align: AlignmentMatrix) -> str:
    if not claude_client:
        return "Examiner synthesis unavailable — ANTHROPIC_API_KEY not set."

    # Build per-chapter + per-subsection summary
    summaries_text = ""
    for cs in chs:
        n_c = sum(1 for c in cs.comments if c.severity=="CRITICAL")
        n_m = sum(1 for c in cs.comments if c.severity=="MODERATE")
        summaries_text += (f"\nCh.{cs.chapter_num} — {cs.chapter_title}: "
                           f"{len(cs.comments)} comments "
                           f"(Critical:{n_c} Moderate:{n_m})\n")
        # Subsection-level breakdown
        for sub in cs.subsections:
            sub_crit = [c for c in sub.comments if c.severity=="CRITICAL"]
            if sub_crit:
                summaries_text += f"  §{sub.subsection_num} {sub.title} (purpose: {sub.expected_purpose[:80]})\n"
                for c in sub_crit[:3]:
                    summaries_text += f"     [CRITICAL] {c.issue}\n"

    gaps_text = "\n".join(f"  • {g}" for g in align.critical_gaps[:8]) or "  (none recorded)"

    template = _EXAMINER_PROMPTS.get(doc_type, _EXAMINER_PROMPTS["MASTERS"])
    prompt = template.format(
        title=spine.title, field=field, institution=inst,
        problem=spine.problem_statement[:400],
        rqs="; ".join(spine.research_questions[:5]) or "NOT FOUND",
        ros="; ".join(spine.research_objectives[:5]) or "NOT FOUND",
        theory=spine.theory_used,
        method=spine.methodology,
        analysis=spine.analysis_technique,
        gt_score=align.golden_thread_score,
        align_verdict=align.overall_verdict,
        gaps=gaps_text,
        summaries=summaries_text[:8000],
        frameworks=CANONICAL_FRAMEWORKS,
        methods=CANONICAL_METHODS,
    )
    try:
        msg = await claude_client.messages.create(
            model="claude-sonnet-4-6", max_tokens=8192, temperature=0.3,
            system=_EXAMINER_SYSTEM,
            messages=[{"role":"user","content":prompt}])
        return msg.content[0].text.strip()
    except Exception as e:
        print(f"Examiner error: {e}")
        return f"Examiner synthesis failed: {e}"


# ── PDF style factory ──────────────────────────────────────────
def pdf_styles():
    def S(n, **k):
        k.setdefault("fontName","Times-Roman")
        k.setdefault("fontSize",11)
        k.setdefault("leading",16)
        k.setdefault("textColor",BLACK)
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
        "Bold":    S("Bold",fontName="Times-Bold",leading=18,spaceAfter=6),
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

def meta_box(S, rows):
    data=[[Paragraph(f'<font name="Helvetica-Bold">{k}</font>',S["Meta"]),
           Paragraph(str(v),S["Meta"])] for k,v in rows]
    t=Table(data,colWidths=[4*cm,W-5*cm-4*cm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),BOX_BG),("BOX",(0,0),(-1,-1),.5,RULE),
        ("INNERGRID",(0,0),(-1,-1),.3,RULE),("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8)]))
    return t

def vbanner(S, text, color, subtitle="Examiner's Recommended Outcome"):
    t=Table([[Paragraph(text.upper(),S["Verdict"])],[Paragraph(subtitle,S["VrdSub"])]],
            colWidths=[W-5*cm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),color),("TOPPADDING",(0,0),(-1,-1),10),
        ("BOTTOMPADDING",(0,0),(-1,-1),10),("LEFTPADDING",(0,0),(-1,-1),16),("RIGHTPADDING",(0,0),(-1,-1),16)]))
    return t

def sev_badge(S, sev, count):
    t=Table([[Paragraph(f"{sev}: {count}",S["Badge"])]],colWidths=[3.5*cm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),SEV_COLORS.get(sev,ACCENT)),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8)]))
    return t

def parse_secs(text):
    pat=re.compile(r'(?:^|\n)\s*(?:\*{0,2})?SECTION\s+\d+\s*[—–\-]+\s*([^\n*]+?)(?:\*{0,2})?\s*\n',re.IGNORECASE)
    ms=list(pat.finditer(text))
    if not ms: return [{"title":"Full Report","body":text.strip()}]
    out=[]
    for i,m in enumerate(ms):
        t=m.group(1).strip().rstrip("*").strip()
        s=m.end(); e=ms[i+1].start() if i+1<len(ms) else len(text)
        out.append({"title":t,"body":text[s:e].strip()})
    return out

def detect_verdict(text):
    u=text.upper()
    for lbl,col in [("READY FOR SUBMISSION",GREEN),("ACCEPT",GREEN),
                    ("PASS WITH MINOR CORRECTIONS",AMBER),("MINOR REVISIONS REQUIRED",AMBER),
                    ("MAJOR REVISIONS REQUIRED",RED),("REFER (RESUBMIT)",RED),
                    ("REJECT AND REWRITE",RED),("REJECT",RED),("RESUBMIT",RED),
                    ("FAIL",RED),("PASS",GREEN)]:
        if lbl in u: return lbl.title(), col
    return "Review Complete", ACCENT

def body2story(S, body):
    items=[]
    for line in body.split("\n"):
        line=line.strip()
        if not line: items.append(Spacer(1,3)); continue
        line=re.sub(r'\*\*(.+?)\*\*',r'<b>\1</b>',line)
        line=re.sub(r'\*(.+?)\*',r'<i>\1</i>',line)
        if re.match(r'^[-•]\s+',line): items.append(Paragraph("• "+line[2:],S["Bullet"]))
        elif re.match(r'^\d+[\.\)]\s+',line): items.append(Paragraph(line,S["Bullet"]))
        else: items.append(Paragraph(line,S["Body"]))
    return items

def type_badge_table(S, text, color):
    t=Table([[Paragraph(text,S["Badge"])]],colWidths=[W-5*cm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),color),("TOPPADDING",(0,0),(-1,-1),7),
        ("BOTTOMPADDING",(0,0),(-1,-1),7),("LEFTPADDING",(0,0),(-1,-1),12),("RIGHTPADDING",(0,0),(-1,-1),12)]))
    return t

def spine_box(S, spine: ThesisSpine):
    """Render the extracted thesis spine as a metadata box."""
    rows = [
        ("Title",             spine.title[:120]),
        ("Discipline",        spine.discipline),
        ("Problem Statement", spine.problem_statement[:300]),
        ("Research Gap",      spine.research_gap[:250]),
        ("Research Questions", "; ".join(spine.research_questions[:6]) or "NOT FOUND"),
        ("Research Objectives","; ".join(spine.research_objectives[:6]) or "NOT FOUND"),
        ("Theory Used",       spine.theory_used),
        ("Variables",         "; ".join(spine.variables[:10]) or "NOT FOUND"),
        ("Methodology",       spine.methodology),
        ("Sampling",          spine.sampling),
        ("Instrument",        spine.instrument),
        ("Analysis Technique",spine.analysis_technique),
        ("Key Findings",      "; ".join(spine.key_findings[:4]) or "NOT FOUND"),
        ("Conclusions",       "; ".join(spine.conclusions[:4]) or "NOT FOUND"),
    ]
    data = [[Paragraph(k, S["SpineK"]), Paragraph(str(v), S["SpineV"])] for k, v in rows]
    t = Table(data, colWidths=[4.5*cm, W-5*cm-4.5*cm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),BOX_BG),
        ("BOX",(0,0),(-1,-1),.5,RULE),("INNERGRID",(0,0),(-1,-1),.3,RULE),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
        ("VALIGN",(0,0),(-1,-1),"TOP")]))
    return t


# ── Output 1: Examiner Report PDF ──────────────────────────────
def build_examiner_pdf(filename, audit_id, doc_type, clf, spine, examiner_text,
                       align: AlignmentMatrix, chs) -> bytes:
    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,leftMargin=2.5*cm,rightMargin=2.5*cm,
                          topMargin=2.2*cm,bottomMargin=2.2*cm)
    S=pdf_styles(); story=[]

    type_label=TYPE_LABELS.get(doc_type,doc_type)
    story.append(type_badge_table(S,type_label.upper(),TYPE_COLORS.get(doc_type,NAVY)))
    story.append(Spacer(1,14))
    story.append(Paragraph("EXAMINER'S AUDIT REPORT",S["Title"]))
    story.append(Paragraph("ThesisSifu Pro v4 — Alignment-Aware Multi-Agent Panel",S["Sub"]))
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

    # NEW: Thesis Spine summary
    story.append(Paragraph("EXTRACTED THESIS SPINE",S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))
    story.append(spine_box(S, spine))
    story.append(Spacer(1,10))

    # Severity badges
    nc=sum(1 for cs in chs for c in cs.comments if c.severity=="CRITICAL")
    nm=sum(1 for cs in chs for c in cs.comments if c.severity=="MODERATE")
    ns=sum(1 for cs in chs for c in cs.comments if c.severity=="SUGGESTION")
    brow=Table([[sev_badge(S,"CRITICAL",nc),Spacer(6,1),
                 sev_badge(S,"MODERATE",nm),Spacer(6,1),
                 sev_badge(S,"SUGGESTION",ns)]],
               colWidths=[3.5*cm,.5*cm,3.5*cm,.5*cm,3.5*cm])
    brow.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))
    story.append(brow); story.append(Spacer(1,10))

    # Golden thread banner (NEW — distinct from verdict)
    gt = align.golden_thread_score
    gt_color = {"STRONG":GREEN,"ACCEPTABLE":AMBER,"WEAK":RED,"BROKEN":RED}.get(gt, ACCENT)
    story.append(vbanner(S, f"Golden Thread: {gt}", gt_color,
                          subtitle="Structural Alignment Score"))
    story.append(Spacer(1,8))

    # Final verdict banner
    vl,vc=detect_verdict(examiner_text)
    story.append(vbanner(S,vl,vc)); story.append(Spacer(1,10)); story.append(HR())

    for sec in parse_secs(examiner_text):
        bl=[Paragraph(sec["title"].upper(),S["SecHead"]),
            HR(c=ACCENT,t=.7,b=0,a=6)]
        bl.extend(body2story(S,sec["body"])); bl.append(Spacer(1,6))
        story.append(KeepTogether(bl[:3])); story.extend(bl[3:])

    # Chapter+subsection summary table
    story.append(PageBreak())
    story.append(Paragraph("CHAPTER + SUBSECTION ISSUE SUMMARY",S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=8))
    td=[[Paragraph(h,S["TblH"]) for h in ["Chapter / Subsection","Critical","Moderate","Suggestions","Total"]]]
    for cs in chs:
        nc2=sum(1 for c in cs.comments if c.severity=="CRITICAL")
        nm2=sum(1 for c in cs.comments if c.severity=="MODERATE")
        ns2=len(cs.comments)-nc2-nm2
        td.append([Paragraph(f"<b>Ch.{cs.chapter_num} — {cs.chapter_title[:35]}</b>",S["TblC"]),
                   Paragraph(str(nc2),S["TblC"]),Paragraph(str(nm2),S["TblC"]),
                   Paragraph(str(ns2),S["TblC"]),Paragraph(str(nc2+nm2+ns2),S["TblC"])])
        for sub in cs.subsections:
            if not sub.comments: continue
            snc=sum(1 for c in sub.comments if c.severity=="CRITICAL")
            snm=sum(1 for c in sub.comments if c.severity=="MODERATE")
            sns=len(sub.comments)-snc-snm
            td.append([Paragraph(f"   §{sub.subsection_num} {sub.title[:40]}",S["TblCSm"]),
                       Paragraph(str(snc),S["TblCSm"]),Paragraph(str(snm),S["TblCSm"]),
                       Paragraph(str(sns),S["TblCSm"]),Paragraph(str(snc+snm+sns),S["TblCSm"])])
    cw=W-5*cm
    t=Table(td,colWidths=[cw*.50,cw*.12,cw*.13,cw*.13,cw*.12])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),("BOX",(0,0),(-1,-1),.5,RULE),
        ("INNERGRID",(0,0),(-1,-1),.3,RULE),("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6)]))
    story.append(t)
    story.append(Spacer(1,20)); story.append(HR())
    story.append(Paragraph("ThesisSifu Pro v4 — Alignment-aware academic audit. "
        "See Report 2 (Annotated Thesis), Report 3 (Commentary Log), "
        "and Report 4 (Alignment Matrix) in this ZIP.",S["Footer"]))

    doc.build(story); buf.seek(0); return buf.getvalue()


# ── Output 4: Alignment Matrix Report PDF (NEW) ────────────────
def build_alignment_pdf(filename, audit_id, doc_type, clf, spine: ThesisSpine,
                        align: AlignmentMatrix, chs) -> bytes:
    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,leftMargin=2*cm,rightMargin=2*cm,
                          topMargin=2.2*cm,bottomMargin=2.2*cm)
    S=pdf_styles(); story=[]

    story.append(type_badge_table(S,"STRUCTURAL ALIGNMENT MATRIX",
                                   TYPE_COLORS.get(doc_type,NAVY)))
    story.append(Spacer(1,14))
    story.append(Paragraph("ALIGNMENT MATRIX REPORT",S["Title"]))
    story.append(Paragraph("ThesisSifu Pro v4 — Golden Thread Audit",S["Sub"]))
    story.append(HR(c=NAVY,t=1.5,b=4,a=10))

    story.append(meta_box(S,[
        ("Document",filename),
        ("Audit ID",audit_id),
        ("Date",datetime.now().strftime("%d %B %Y, %H:%M")),
        ("Report","4 of 4 — Structural Alignment Matrix"),
    ]))
    story.append(Spacer(1,10))

    # Golden thread banner
    gt = align.golden_thread_score
    gt_color = {"STRONG":GREEN,"ACCEPTABLE":AMBER,"WEAK":RED,"BROKEN":RED}.get(gt, ACCENT)
    story.append(vbanner(S, f"Golden Thread: {gt}", gt_color,
                          subtitle="Problem → RQ → RO → Method → Analysis → Finding → Conclusion"))
    story.append(Spacer(1,10))

    # Overall verdict
    story.append(Paragraph("OVERALL ALIGNMENT VERDICT", S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))
    story.append(Paragraph(align.overall_verdict or "(no verdict generated)", S["Body"]))
    story.append(Spacer(1,8))

    # Spine reminder
    story.append(Paragraph("THESIS SPINE (for reference)", S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))
    story.append(spine_box(S, spine))
    story.append(Spacer(1,10))

    # Alignment matrix
    story.append(PageBreak())
    story.append(Paragraph("ALIGNMENT MATRIX — RQ-LEVEL TRACE", S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))
    story.append(Paragraph(
        "Each row traces one research question through the thesis. "
        "<b>Status</b> indicates whether the chain holds together.",
        S["Body"]))
    story.append(Spacer(1,6))

    if not align.rows:
        story.append(Paragraph("<i>No alignment rows produced — likely insufficient spine data.</i>",
                                S["Body"]))
    else:
        headers = ["RQ", "RO / H", "Method + Analysis", "Finding", "Conclusion", "Status"]
        td = [[Paragraph(f"<b>{h}</b>", S["TblH"]) for h in headers]]
        for r in align.rows:
            status_color = ALIGN_COLORS.get(r.status, ACCENT)
            status_para = Paragraph(
                f'<font color="{status_color.hexval()}"><b>{r.status}</b></font>',
                S["TblCSm"])
            roh = r.ro + (f" / {r.hypothesis}" if r.hypothesis and r.hypothesis != "—" else "")
            ma  = f"{r.method}<br/><i>{r.analysis}</i>"
            td.append([
                Paragraph(r.rq[:120], S["TblCSm"]),
                Paragraph(roh[:100], S["TblCSm"]),
                Paragraph(ma[:140], S["TblCSm"]),
                Paragraph(r.finding[:120], S["TblCSm"]),
                Paragraph(r.conclusion[:120], S["TblCSm"]),
                status_para,
            ])
        cw = W - 4*cm
        t = Table(td, colWidths=[cw*.18, cw*.16, cw*.20, cw*.17, cw*.17, cw*.12])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),
            ("BOX",(0,0),(-1,-1),.5,RULE),("INNERGRID",(0,0),(-1,-1),.3,RULE),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
            ("VALIGN",(0,0),(-1,-1),"TOP")]))
        story.append(t)
        story.append(Spacer(1,10))

        # Row notes
        story.append(Paragraph("ROW-BY-ROW ALIGNMENT NOTES", S["SecHead"]))
        story.append(HR(c=NAVY,t=1,b=0,a=6))
        for i, r in enumerate(align.rows, 1):
            status_color = ALIGN_COLORS.get(r.status, ACCENT)
            badge = Table([[Paragraph(r.status, S["Badge"])]], colWidths=[3*cm])
            badge.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),status_color),
                ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3)]))
            hdr_row = Table([[
                Paragraph(f"<b>RQ{i}.</b> {r.rq[:140]}", S["Body"]),
                badge,
            ]], colWidths=[W-5*cm-3.2*cm, 3*cm])
            hdr_row.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),
                ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))
            story.append(hdr_row)
            if r.note:
                story.append(Paragraph(f"<i>{r.note}</i>", S["Excerpt"]))
            story.append(Spacer(1,6))

    # Critical gaps
    if align.critical_gaps:
        story.append(PageBreak())
        story.append(Paragraph("CRITICAL STRUCTURAL GAPS", S["SecHead"]))
        story.append(HR(c=NAVY,t=1,b=0,a=6))
        for g in align.critical_gaps:
            story.append(Paragraph(f"• {g}", S["Bullet"]))
        story.append(Spacer(1,10))

    # Structural recommendations
    if align.structural_recommendations:
        story.append(Paragraph("STRUCTURAL RECOMMENDATIONS", S["SecHead"]))
        story.append(HR(c=NAVY,t=1,b=0,a=6))
        for rec in align.structural_recommendations:
            story.append(Paragraph(f"• {rec}", S["Bullet"]))
        story.append(Spacer(1,10))

    # Chapter delivery scorecard — does each chapter deliver on its expected purpose?
    story.append(PageBreak())
    story.append(Paragraph("CHAPTER DELIVERY SCORECARD", S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=6))
    story.append(Paragraph(
        "Did each chapter/subsection deliver on its <b>expected purpose</b>? "
        "A high count of critical issues against a subsection signals it has not.",
        S["Body"]))
    story.append(Spacer(1,6))

    headers = ["Chapter / Subsection", "Expected Purpose", "Issues (C / M / S)"]
    td = [[Paragraph(f"<b>{h}</b>", S["TblH"]) for h in headers]]
    for cs in chs:
        td.append([
            Paragraph(f"<b>Ch.{cs.chapter_num} — {cs.chapter_title[:50]}</b>", S["TblC"]),
            Paragraph("(see subsections below)", S["TblCSm"]),
            Paragraph(f"{sum(1 for c in cs.comments if c.severity=='CRITICAL')} / "
                       f"{sum(1 for c in cs.comments if c.severity=='MODERATE')} / "
                       f"{sum(1 for c in cs.comments if c.severity=='SUGGESTION')}", S["TblC"]),
        ])
        for sub in cs.subsections:
            snc=sum(1 for c in sub.comments if c.severity=="CRITICAL")
            snm=sum(1 for c in sub.comments if c.severity=="MODERATE")
            sns=len(sub.comments)-snc-snm
            td.append([
                Paragraph(f"   §{sub.subsection_num} {sub.title[:50]}", S["TblCSm"]),
                Paragraph(sub.expected_purpose[:120], S["TblCSm"]),
                Paragraph(f"{snc} / {snm} / {sns}", S["TblCSm"]),
            ])
    cw = W - 4*cm
    t = Table(td, colWidths=[cw*.34, cw*.50, cw*.16])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),
        ("BOX",(0,0),(-1,-1),.5,RULE),("INNERGRID",(0,0),(-1,-1),.3,RULE),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(t)
    story.append(Spacer(1,16)); story.append(HR())

    story.append(Paragraph(
        "ThesisSifu Pro v4 — Structural Alignment Matrix. "
        "Use alongside Report 1 (Examiner Audit) and Report 3 (Commentary Log) "
        "to triangulate revision priorities.",
        S["Footer"]))

    doc.build(story); buf.seek(0); return buf.getvalue()


# ── Output 2: Annotated DOCX ───────────────────────────────────
def _get_or_create_comments_part(doc):
    comments_reltype = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
    for rel in doc.part.rels.values():
        if rel.reltype == comments_reltype:
            return rel.target_part
    xml_str = '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    element = parse_xml(xml_str)
    uri = PackURI('/word/comments.xml')
    ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    cp = XmlPart(uri, ct, element, doc.part.package)
    doc.part.relate_to(cp, comments_reltype)
    return cp


def _insert_comment(doc, para, cid, author, text, date_str):
    cp = _get_or_create_comments_part(doc)
    comments_root = cp.element

    comment_el = OxmlElement('w:comment')
    comment_el.set(qn('w:id'), str(cid))
    comment_el.set(qn('w:author'), author)
    comment_el.set(qn('w:date'), date_str)
    comment_el.set(qn('w:initials'), 'TS')

    p_el = OxmlElement('w:p')
    r_el = OxmlElement('w:r')
    t_el = OxmlElement('w:t')
    t_el.text = text
    t_el.set(qn('xml:space'), 'preserve')
    r_el.append(t_el); p_el.append(r_el); comment_el.append(p_el)
    comments_root.append(comment_el)

    px = para._p
    crs = OxmlElement('w:commentRangeStart')
    crs.set(qn('w:id'), str(cid)); px.insert(0, crs)
    cre = OxmlElement('w:commentRangeEnd')
    cre.set(qn('w:id'), str(cid)); px.append(cre)
    rr  = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rs  = OxmlElement('w:rStyle')
    rs.set(qn('w:val'), 'CommentReference')
    rpr.append(rs); rr.append(rpr)
    cr  = OxmlElement('w:commentReference')
    cr.set(qn('w:id'), str(cid)); rr.append(cr); px.append(rr)


def _comment_text(c: ParagraphComment) -> str:
    """Format a comment for inline Word/PDF display — now includes subsection +
    suggested framework/method fields."""
    sev = {"CRITICAL":"[CRITICAL]","MODERATE":"[MODERATE]",
           "SUGGESTION":"[SUGGESTION]"}.get(c.severity,"[NOTE]")
    fw  = f"\n\nSUGGESTED FRAMEWORK: {c.suggested_framework}" if c.suggested_framework else ""
    mt  = f"\n\nSUGGESTED METHOD/REFERENCE: {c.suggested_method}" if c.suggested_method else ""
    return (f"{sev} ThesisSifu Pro v4 — §{c.subsection} {c.subsection_title}\n\n"
            f"ISSUE: {c.issue}\n\n"
            f"RECOMMENDATION: {c.recommendation}\n\n"
            f"LITERATURE NEEDED: {c.literature_needed}\n\n"
            f"THEORY/FRAMEWORK: {c.theory_needed}"
            f"{fw}{mt}")


def build_annotated_docx(content, filename, audit_id, chs, clf) -> bytes:
    doc = DocxDocument(io.BytesIO(content))
    date_str = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

    cp = _get_or_create_comments_part(doc)
    existing_ids = []
    for comment in cp.element.xpath('.//w:comment'):
        c_id = comment.get(qn('w:id'))
        if c_id is not None:
            existing_ids.append(int(c_id))
    cid = max(existing_ids) + 1 if existing_ids else 0

    lookup: dict[str, list] = {}
    for cs in chs:
        for c in cs.comments:
            key = c.para_excerpt[:50].lower().strip()
            lookup.setdefault(key, []).append(c)

    matched = set()
    for para in doc.paragraphs:
        if not para.text.strip() or len(para.text.strip()) < 30:
            continue
        pkey = para.text[:50].lower().strip()
        best, best_sc = None, 0
        for ekey, cmts in lookup.items():
            aw = set(pkey.split()); bw = set(ekey.split())
            sc = len(aw & bw) / max(len(aw), 1)
            if sc > best_sc and sc > 0.35:
                best_sc = sc; best = (ekey, cmts)
        if best:
            for c in best[1]:
                if id(c) in matched: continue
                matched.add(id(c))
                try:
                    _insert_comment(doc, para, cid, "ThesisSifu Pro v4",
                                    _comment_text(c), date_str)
                    cid += 1
                except Exception as e:
                    print(f"Comment insert error {cid}: {e}")
                    cid += 1

    hdr = doc.sections[0].header
    if hdr.paragraphs:
        hdr.paragraphs[0].text = (
            f"ThesisSifu Pro v4 | Annotated Thesis | Audit ID: {audit_id} | "
            f"{datetime.now().strftime('%d %B %Y')}"
        )
    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf.getvalue()


def build_annotated_pdf(content: bytes, filename: str, audit_id: str, chs: list) -> bytes:
    doc = fitz.open(stream=content, filetype="pdf")
    for cs in chs:
        for c in cs.comments:
            search_text = c.para_excerpt[:35].replace('\n', ' ').strip()
            start_page = max(0, c.page_estimate - 2)
            end_page = min(len(doc), c.page_estimate + 2)
            annotated = False
            txt = _comment_text(c)

            for p_num in range(start_page, end_page):
                page = doc[p_num]
                rects = page.search_for(search_text)
                if rects:
                    rect = rects[0]
                    point = fitz.Point(rect.x0 - 15, rect.y0)
                    annot = page.add_text_annot(point, txt)
                    annot.set_info(title="ThesisSifu Pro v4", content=txt)
                    if c.severity == "CRITICAL":   annot.set_colors(stroke=(1, 0, 0))
                    elif c.severity == "MODERATE": annot.set_colors(stroke=(1, 0.5, 0))
                    else:                           annot.set_colors(stroke=(0, 0, 1))
                    annot.update(); annotated = True; break

            if not annotated:
                fallback_page = min(c.page_estimate - 1, len(doc) - 1)
                page = doc[max(0, fallback_page)]
                point = fitz.Point(30, 30)
                annot = page.add_text_annot(point, f"[TEXT NOT FOUND ON PAGE]\n{txt}")
                annot.set_colors(stroke=(0.5, 0.5, 0.5)); annot.update()
    return doc.write()


# ── Output 3: Commentary Report PDF (subsection-grouped) ──────
def build_commentary_pdf(filename, audit_id, doc_type, clf, spine, chs) -> bytes:
    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,leftMargin=2.5*cm,rightMargin=2.5*cm,
                          topMargin=2.2*cm,bottomMargin=2.2*cm)
    S=pdf_styles(); story=[]

    story.append(type_badge_table(S,"PARAGRAPH-LEVEL COMMENTARY REPORT",
                                  TYPE_COLORS.get(doc_type,NAVY)))
    story.append(Spacer(1,14))
    story.append(Paragraph("COMMENTARY REPORT",S["Title"]))
    story.append(Paragraph("ThesisSifu Pro v4 — Subsection-Level Supervisor Notes",S["Sub"]))
    story.append(HR(c=NAVY,t=1.5,b=4,a=10))

    au=clf.get("authors","UNKNOWN")
    story.append(meta_box(S,[
        ("Document",filename),("Type",TYPE_LABELS.get(doc_type,doc_type)),
        ("Title",(spine.title[:85]+"…") if len(spine.title)>85 else spine.title),
        ("Author(s)",(au[:80]+"…") if len(au)>80 else au),
        ("Audit ID",audit_id),("Date",datetime.now().strftime("%d %B %Y, %H:%M")),
        ("Report","3 of 4 — Paragraph Commentary Log"),
    ]))
    story.append(Spacer(1,10))

    all_c=[c for cs in chs for c in cs.comments]
    nc=sum(1 for c in all_c if c.severity=="CRITICAL")
    nm=sum(1 for c in all_c if c.severity=="MODERATE")
    ns=sum(1 for c in all_c if c.severity=="SUGGESTION")

    story.append(Paragraph("OVERALL COMMENT DISTRIBUTION",S["SecHead"]))
    cw=W-5*cm
    td=[[Paragraph(h,S["TblH"]) for h in ["Severity","Count","Description"]],
        [Paragraph("CRITICAL",S["TblC"]),Paragraph(str(nc),S["TblC"]),
         Paragraph("Fundamental flaws requiring resolution",S["TblC"])],
        [Paragraph("MODERATE",S["TblC"]),Paragraph(str(nm),S["TblC"]),
         Paragraph("Significant weaknesses requiring revision",S["TblC"])],
        [Paragraph("SUGGESTION",S["TblC"]),Paragraph(str(ns),S["TblC"]),
         Paragraph("Improvement opportunities",S["TblC"])],
        [Paragraph("TOTAL",S["TblC"]),Paragraph(str(len(all_c)),S["TblC"]),
         Paragraph("",S["TblC"])]]
    t=Table(td,colWidths=[cw*.22,cw*.12,cw*.66])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),("BOX",(0,0),(-1,-1),.5,RULE),
        ("INNERGRID",(0,0),(-1,-1),.3,RULE),("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7)]))
    story.append(t); story.append(Spacer(1,16)); story.append(HR())

    for cs in chs:
        if not cs.comments: continue
        story.append(Paragraph(f"CHAPTER {cs.chapter_num} — {cs.chapter_title.upper()}",S["ChHead"]))
        nc2=sum(1 for c in cs.comments if c.severity=="CRITICAL")
        nm2=sum(1 for c in cs.comments if c.severity=="MODERATE")
        story.append(Paragraph(
            f"<b>{len(cs.comments)} comments</b> — "
            f"Critical: {nc2} | Moderate: {nm2} | Suggestions: {len(cs.comments)-nc2-nm2}",
            S["Body"]))
        story.append(HR(c=RULE,t=.4,b=2,a=6))

        # Group by subsection
        for sub in cs.subsections:
            if not sub.comments: continue
            story.append(Paragraph(
                f"§{sub.subsection_num} — {sub.title}", S["SubHead"]))
            story.append(Paragraph(
                f"<i>Expected purpose: {sub.expected_purpose}</i>", S["Excerpt"]))
            story.append(Spacer(1,4))

            for i, c in enumerate(sub.comments, 1):
                sc = SEV_COLORS.get(c.severity, ACCENT)

                hdr_row = Table(
                    [[Paragraph(c.severity,S["Badge"]),Spacer(6,1),
                      Paragraph(f"§{sub.subsection_num} Comment {i} | Page ~{c.page_estimate}",S["Meta"])]],
                    colWidths=[2.2*cm,.4*cm,W-5*cm-2.6*cm])
                hdr_row.setStyle(TableStyle([("BACKGROUND",(0,0),(0,0),sc),
                    ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
                    ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE")]))

                excerpt_box = Table(
                    [[Paragraph(f'<i>"{c.para_excerpt[:120]}..."</i>',S["Excerpt"])]],
                    colWidths=[W-5*cm])
                excerpt_box.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),LGRAY),
                    ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
                    ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8)]))

                fw_line = f'<br/><br/><b>Suggested framework:</b> {c.suggested_framework}' if c.suggested_framework else ""
                mt_line = f'<br/><br/><b>Suggested method/reference:</b> {c.suggested_method}' if c.suggested_method else ""

                detail_box = Table(
                    [[Paragraph(
                        f'<b>Issue:</b> {c.issue}<br/><br/>'
                        f'<b>Recommendation:</b> {c.recommendation}<br/><br/>'
                        f'<b>Literature needed:</b> {c.literature_needed}<br/><br/>'
                        f'<b>Theory / Framework:</b> {c.theory_needed}'
                        f'{fw_line}{mt_line}',
                        S["Detail"])]],
                    colWidths=[W-5*cm])
                detail_box.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),BOX_BG),
                    ("BOX",(0,0),(-1,-1),.4,RULE),("TOPPADDING",(0,0),(-1,-1),7),
                    ("BOTTOMPADDING",(0,0),(-1,-1),7),("LEFTPADDING",(0,0),(-1,-1),10),
                    ("RIGHTPADDING",(0,0),(-1,-1),10)]))

                block=[hdr_row,excerpt_box,detail_box,Spacer(1,8)]
                story.append(KeepTogether(block[:2])); story.extend(block[2:])

        story.append(Spacer(1,12)); story.append(HR())

    story.append(Spacer(1,6))
    story.append(Paragraph(
        "ThesisSifu Pro v4 — Subsection-level paragraph commentary log. "
        "See Report 1 (Examiner Audit) for holistic verdict, "
        "Report 2 (Annotated Thesis) for inline comments, "
        "and Report 4 (Alignment Matrix) for structural verification.",S["Footer"]))

    doc.build(story); buf.seek(0); return buf.getvalue()


# ── Main endpoint ──────────────────────────────────────────────
@app.post("/audit")
async def audit_document(file: UploadFile = File(...)):
    content  = await file.read()
    filename = file.filename or "document"

    if not filename.lower().endswith((".pdf",".docx")):
        raise HTTPException(400, "Only PDF and DOCX files are supported.")

    audit_id = "SUP-" + hashlib.md5(content).hexdigest()[:10].upper()

    # 1. Extract text
    full_text = extract_text(content, filename)
    if not full_text.strip():
        full_text = "[Document appears image-based or empty]"

    # 2. Stage 0: spine extraction + Stage 1: classification (parallel)
    spine_task = asyncio.create_task(extract_spine(full_text))
    clf_task   = asyncio.create_task(classify_document(full_text))
    spine, clf = await asyncio.gather(spine_task, clf_task)

    doc_type = clf.get("type","MASTERS")
    field    = clf.get("field","UNKNOWN")
    inst     = clf.get("institution","UNKNOWN")
    # If classifier got a title but spine didn't, fill it in
    if spine.title == "UNKNOWN" and clf.get("title","UNKNOWN") != "UNKNOWN":
        spine.title = clf["title"]

    # 3. Split chapters + subsections
    chs = split_chapters(full_text)
    for cs in chs:
        cs.subsections = split_subsections(cs)

    # 4. Stage 2: subsection-level paragraph audit (parallel, max 6 concurrent)
    sem = asyncio.Semaphore(6)
    async def bounded(sub: Subsection, ch_title: str):
        async with sem:
            return sub, await audit_subsection(sub, ch_title, doc_type, spine)

    tasks = []
    for cs in chs:
        for sub in cs.subsections:
            tasks.append(bounded(sub, cs.chapter_title))
    results = await asyncio.gather(*tasks)

    # Attach comments to subsections AND flatten to chapter level
    for sub, cmts in results:
        sub.comments = cmts
    for cs in chs:
        cs.comments = [c for sub in cs.subsections for c in sub.comments]

    # 5. Stage 3: alignment audit (Claude)
    align = await run_alignment_audit(spine, chs)

    # 6. Stage 4: holistic examiner synthesis
    examiner_text = await run_examiner(doc_type, spine, field, inst, chs, align)

    # 7. Build four outputs
    pdf1 = build_examiner_pdf(filename, audit_id, doc_type, clf, spine,
                              examiner_text, align, chs)

    if filename.lower().endswith(".pdf"):
        output2 = build_annotated_pdf(content, filename, audit_id, chs)
        out2_name = "2_Annotated_Thesis.pdf"
    else:
        output2 = build_annotated_docx(content, filename, audit_id, chs, clf)
        out2_name = "2_Annotated_Thesis.docx"

    pdf3 = build_commentary_pdf(filename, audit_id, doc_type, clf, spine, chs)
    pdf4 = build_alignment_pdf(filename, audit_id, doc_type, clf, spine, align, chs)

    # 8. ZIP and return
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("1_Examiner_Audit_Report.pdf", pdf1)
        zf.writestr(out2_name, output2)
        zf.writestr("3_Commentary_Report.pdf", pdf3)
        zf.writestr("4_Alignment_Matrix_Report.pdf", pdf4)
    tmp.close()

    return FileResponse(
        path=tmp.name, media_type="application/zip",
        filename=f"ThesisSifu_v4_Audit_{audit_id}.zip",
        background=BackgroundTask(lambda: os.unlink(tmp.name)),
        headers={"Access-Control-Expose-Headers":"Content-Disposition"},
    )


@app.get("/health")
async def health():
    return {
        "status":                "ok",
        "classifier_available":  gemini_client is not None,
        "audit_model_available": claude_client is not None,
        "version":               "4.0.0",
        "stages":                ["spine_extraction", "chapter_subsection_split",
                                  "subsection_paragraph_audit", "alignment_audit",
                                  "holistic_examiner"],
        "outputs":               ["1_Examiner_Audit_Report.pdf",
                                  "2_Annotated_Thesis.(docx|pdf)",
                                  "3_Commentary_Report.pdf",
                                  "4_Alignment_Matrix_Report.pdf"],
    }
