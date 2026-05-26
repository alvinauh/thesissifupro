"""
PhD Supervisor Academic Audit System
=====================================
Classifies document type, then produces a structured high-level
internal-examiner-quality audit report as a downloadable PDF.

Document types detected:
  - Scopus/journal article
  - Undergraduate thesis / FYP
  - Master's thesis
  - PhD dissertation

Endpoint:  POST /audit
           multipart/form-data  { file: <pdf|docx> }

Returns:   application/zip  { Audit_Report.pdf, Summary.txt }
"""

import io
import os
import re
import asyncio
import zipfile
import hashlib
import tempfile
from datetime import datetime

import pypdf
from docx import Document as DocxDocument
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from starlette.background import BackgroundTask

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY, TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, HRFlowable,
    Table, TableStyle, PageBreak, KeepTogether, ListFlowable, ListItem
)

import google.generativeai as genai

# ─── Config ───────────────────────────────────────────────────
app = FastAPI(title="PhD Supervisor Audit System")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

api_key = os.environ.get("GEMINI_API_KEY")
if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")
else:
    model = None
    print("WARNING: GEMINI_API_KEY not set — AI calls will be skipped.")

W, H = A4
BLACK   = colors.HexColor("#111111")
NAVY    = colors.HexColor("#0C2340")
ACCENT  = colors.HexColor("#1A4A7A")
RULE    = colors.HexColor("#AAAAAA")
BOX_BG  = colors.HexColor("#F4F6F9")
RED     = colors.HexColor("#A30000")
AMBER   = colors.HexColor("#7A4500")
GREEN   = colors.HexColor("#0A5C2A")
LGRAY   = colors.HexColor("#EEEEEE")


# ═══════════════════════════════════════════════════════════════
#  TEXT EXTRACTION
# ═══════════════════════════════════════════════════════════════
def extract_text(content: bytes, filename: str) -> str:
    fname = filename.lower()
    if fname.endswith(".pdf"):
        reader = pypdf.PdfReader(io.BytesIO(content))
        parts = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                parts.append(t)
        return "\n".join(parts)
    elif fname.endswith(".docx"):
        doc = DocxDocument(io.BytesIO(content))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    return ""


# ═══════════════════════════════════════════════════════════════
#  DOCUMENT CLASSIFIER
# ═══════════════════════════════════════════════════════════════
CLASSIFICATION_PROMPT = """
You are an expert academic classification system.

Read the following text excerpt (beginning of an academic document) and determine EXACTLY what type of document it is.

Choose ONLY ONE of these labels:
- JOURNAL_ARTICLE   : A paper submitted to or published in a Scopus-indexed or peer-reviewed journal
- UNDERGRADUATE     : An undergraduate final year project (FYP), capstone, or bachelor's thesis
- MASTERS           : A master's thesis (by research or coursework-research)
- PHD               : A PhD dissertation or doctoral thesis

Then provide:
1. The label (one of the four above)
2. Confidence: HIGH / MEDIUM / LOW
3. Key signals that led to your classification (2-3 bullet points)
4. The document's apparent title
5. The apparent author(s)
6. The apparent field/discipline

Respond in EXACTLY this JSON format (no markdown fences):
{
  "type": "JOURNAL_ARTICLE",
  "confidence": "HIGH",
  "signals": ["signal 1", "signal 2", "signal 3"],
  "title": "detected title or UNKNOWN",
  "authors": "detected authors or UNKNOWN",
  "field": "detected field or UNKNOWN"
}

Document excerpt:
\"\"\"
{text}
\"\"\"
"""

async def classify_document(text: str) -> dict:
    default = {
        "type": "MASTERS",
        "confidence": "LOW",
        "signals": ["Could not classify — defaulting to Master's thesis"],
        "title": "UNKNOWN",
        "authors": "UNKNOWN",
        "field": "UNKNOWN"
    }
    if not model or not text.strip():
        return default
    try:
        excerpt = text[:6000]
        prompt = CLASSIFICATION_PROMPT.format(text=excerpt)
        response = await model.generate_content_async(prompt)
        raw = response.text.strip()
        # Strip markdown fences if present
        raw = re.sub(r"```(?:json)?", "", raw).strip().rstrip("`").strip()
        import json
        result = json.loads(raw)
        return result
    except Exception as e:
        print(f"Classification error: {e}")
        return default


# ═══════════════════════════════════════════════════════════════
#  AUDIT PROMPT TEMPLATES  (one per document type)
# ═══════════════════════════════════════════════════════════════
AUDIT_SYSTEM = """
You are a senior academic reviewer and PhD supervisor with 25 years of experience.
Your role is to produce a HIGH-LEVEL, RIGOROUS, CRITICAL audit of an academic document.

CRITICAL RULES:
1. You must respond in the SAME LANGUAGE as the document text.
2. You are NOT a grammar checker — you are an intellectual critic.
3. Be specific: cite page numbers, section numbers, or direct quotes where possible.
4. Be honest: if something is fundamentally flawed, say so clearly.
5. Structure your response EXACTLY as specified in the prompt.
6. Every finding must cite where in the document the evidence is.
"""

JOURNAL_PROMPT = """
You are reviewing a Scopus/peer-reviewed journal article submission.

Produce a structured audit with the following EXACT sections. For each section write 3-6 sentences of substantive critique.

SECTION 1 — CONTRIBUTION & NOVELTY
Does the paper make a genuine, original contribution to the field? Is the research gap clearly identified and justified? Is the novelty claim supported by the literature review?

SECTION 2 — LITERATURE REVIEW & CITATION INTEGRITY
Is the literature current (within 5 years for fast-moving fields)? Are foundational works cited? Are there missing key references? Identify any citation errors or inconsistencies.

SECTION 3 — RESEARCH DESIGN & METHODOLOGY
Is the chosen methodology appropriate and justified? Is the research design rigorous and replicable? Are limitations of the design acknowledged?

SECTION 4 — DATA, ANALYSIS & RESULTS
Are the statistical/analytical methods appropriate? Are results clearly reported with effect sizes and confidence intervals where applicable? Are tables and figures properly labelled and interpreted?

SECTION 5 — ARGUMENT COHERENCE & LOGICAL FLOW
Does the paper build a coherent argument from introduction to conclusion? Are research questions/hypotheses explicitly stated and fully answered? Is there logical consistency across sections?

SECTION 6 — DISCUSSION & CONCLUSION
Does the discussion engage critically with the results or merely restate them? Are claims proportional to the evidence? Does the conclusion overstate findings?

SECTION 7 — SCOPUS PUBLICATION READINESS
Would this paper be accepted by a Scopus Q1/Q2 journal in its current state? What are the TOP 3 critical revisions required before submission?

SECTION 8 — CITATION & REFERENCE AUDIT
List specific citation errors found (date mismatches, missing references, format inconsistencies).

SECTION 9 — EXAMINER'S VERDICT
One of: READY FOR SUBMISSION / MINOR REVISIONS REQUIRED / MAJOR REVISIONS REQUIRED / REJECT AND REWRITE
Justify in 2-3 sentences.

Document text (up to 20,000 chars):
\"\"\"
{text}
\"\"\"
"""

UNDERGRADUATE_PROMPT = """
You are reviewing an undergraduate final year project (FYP) or bachelor's thesis.

Produce a structured audit with the following EXACT sections. For each section write 3-5 sentences of substantive critique.

SECTION 1 — RESEARCH FOCUS & SCOPE APPROPRIATENESS
Is the scope appropriate for an undergraduate project? Is the research question achievable within typical undergraduate constraints? Does the student show understanding of the field?

SECTION 2 — LITERATURE REVIEW
Is the literature review sufficient for undergraduate level? Are key concepts explained accurately? Are there major gaps in the theoretical grounding?

SECTION 3 — METHODOLOGY
Is the chosen method appropriate and correctly applied? Has the student justified their methodological choices? Are ethical considerations addressed?

SECTION 4 — DATA COLLECTION & ANALYSIS
Is the data collection method appropriate? Is the analysis correctly applied? Are findings clearly presented?

SECTION 5 — CRITICAL THINKING & ORIGINALITY
Does the student demonstrate independent thinking, or is the work mostly descriptive? Are findings interpreted with appropriate academic depth?

SECTION 6 — WRITING QUALITY & ACADEMIC CONVENTIONS
Is the writing clear and academically appropriate? Are citations formatted consistently? Is the structure logical?

SECTION 7 — MAJOR ISSUES REQUIRING ATTENTION
List the top issues the student must address before submission.

SECTION 8 — SUPERVISOR'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS / FAIL — RESUBMIT
Justify in 2-3 sentences.

Document text (up to 20,000 chars):
\"\"\"
{text}
\"\"\"
"""

MASTERS_PROMPT = """
You are the internal reader reviewing a Master's thesis (by research).
Act as a critical examiner who will sit on the viva panel.

Produce a structured chapter-by-chapter audit with the following EXACT sections.

SECTION 1 — THESIS ALIGNMENT & GOLDEN THREAD
Does the thesis maintain a consistent argument from title to conclusion? Are the research objectives, research questions, and hypotheses perfectly aligned? Is there a clear "golden thread"?

SECTION 2 — LITERATURE REVIEW QUALITY
Is the literature review critical or merely descriptive? Are foundational and recent sources balanced? Are seminal works in the field present? Identify specific citation problems.

SECTION 3 — THEORETICAL FRAMEWORK
Is the theoretical framework appropriate, coherent, and consistently applied throughout the thesis? Are theories applied correctly to the study context?

SECTION 4 — RESEARCH DESIGN & METHODOLOGY
Is the chosen methodology justified and appropriate? Are sampling decisions explained and defensible? Are threats to validity acknowledged and mitigated?

SECTION 5 — DATA ROBUSTNESS & INSTRUMENTATION
Are the instruments valid, reliable, and appropriate for the sample? Is the data collection process rigorous? Are instrument modifications justified?

SECTION 6 — FINDINGS & ANALYSIS
Are the statistical/analytical methods appropriate? Are assumption tests reported? Are effect sizes reported? Are findings clearly and accurately presented?

SECTION 7 — DISCUSSION & ARGUMENT STRENGTH
Does the candidate engage critically with findings or merely describe them? Are claims proportional to evidence? Is the discussion well-referenced?

SECTION 8 — AUTHENTICITY & RESEARCHER POSITIONALITY
Does the work demonstrate original intellectual contribution? Is the researcher's positionality acknowledged where relevant? Are there any authenticity concerns?

SECTION 9 — CITATION & REFERENCE INTEGRITY
List all citation errors found: date mismatches, missing references, references in list not cited in text, format inconsistencies.

SECTION 10 — CRITICAL VIVA QUESTIONS
List 5 specific viva questions the candidate must be prepared to defend.

SECTION 11 — EXAMINER'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS REQUIRED / RESUBMIT
Justify fully in 4-5 sentences. List the TOP 5 critical issues that must be resolved.

Document text (up to 20,000 chars):
\"\"\"
{text}
\"\"\"
"""

PHD_PROMPT = """
You are an external examiner reviewing a PhD dissertation.
This is the highest level of academic scrutiny. Be rigorous, specific, and uncompromising.

Produce a structured audit with the following EXACT sections.

SECTION 1 — ORIGINAL CONTRIBUTION TO KNOWLEDGE
Does the dissertation make a clear, original, and significant contribution to the field? Is the contribution explicitly stated and consistently demonstrated? Is the claim to novelty defensible?

SECTION 2 — MASTERY OF THE LITERATURE
Does the candidate demonstrate comprehensive, critical command of the field's literature? Are competing theoretical positions fairly represented? Are the most recent and seminal works engaged with?

SECTION 3 — THEORETICAL & CONCEPTUAL FRAMEWORK
Is the theoretical framework sophisticated, coherent, and fit for purpose? Is it consistently applied throughout the dissertation? Are theoretical assumptions made explicit?

SECTION 4 — RESEARCH DESIGN & EPISTEMOLOGICAL ALIGNMENT
Is the research paradigm clearly articulated? Is the methodology epistemologically consistent with the research approach? Are design choices fully justified with reference to the literature?

SECTION 5 — METHODOLOGICAL RIGOUR
Is the research methodology executed with PhD-level rigour? Are all threats to validity/trustworthiness identified and addressed? Is the study replicable from the methodology chapter alone?

SECTION 6 — DATA QUALITY & ANALYTICAL DEPTH
Is the analysis PhD-level in its depth and sophistication? Are analytical methods appropriate and correctly applied? Are findings interpreted with theoretical insight rather than mere description?

SECTION 7 — ARGUMENT ARCHITECTURE & INTELLECTUAL COHERENCE
Does the dissertation build a sustained, coherent intellectual argument? Are all chapters necessary and logically connected? Is there any contradictory reasoning across chapters?

SECTION 8 — DISCUSSION & THEORETICAL CONTRIBUTION
Does the discussion advance theory, not just report findings? Does the candidate position their work within broader scholarly debates? Are limitations treated with intellectual honesty?

SECTION 9 — WRITING & SCHOLARLY VOICE
Is the writing at the level expected of a PhD? Is the scholarly voice authoritative and consistent? Are technical concepts handled with precision?

SECTION 10 — CITATION & REFERENCE INTEGRITY
List all citation errors: date mismatches, missing references, uncited works in reference list, inconsistent formats.

SECTION 11 — CRITICAL EXAMINATION QUESTIONS
List 8 searching questions the candidate must be prepared to defend in the viva voce.

SECTION 12 — EXAMINER'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS / REFER (RESUBMIT) / FAIL
Provide a full examiner's statement (6-8 sentences) covering: the significance of the contribution, the severity of the issues found, and the specific conditions that must be met before the degree can be awarded.

Document text (up to 20,000 chars):
\"\"\"
{text}
\"\"\"
"""

PROMPT_MAP = {
    "JOURNAL_ARTICLE": JOURNAL_PROMPT,
    "UNDERGRADUATE":   UNDERGRADUATE_PROMPT,
    "MASTERS":         MASTERS_PROMPT,
    "PHD":             PHD_PROMPT,
}

TYPE_LABELS = {
    "JOURNAL_ARTICLE": "Scopus / Peer-Reviewed Journal Article",
    "UNDERGRADUATE":   "Undergraduate Thesis / Final Year Project",
    "MASTERS":         "Master's Thesis (By Research)",
    "PHD":             "PhD Dissertation",
}

VERDICT_COLORS = {
    "PASS":                    GREEN,
    "READY FOR SUBMISSION":    GREEN,
    "MINOR REVISIONS":         AMBER,
    "PASS WITH MINOR":         AMBER,
    "MAJOR REVISIONS":         RED,
    "REJECT":                  RED,
    "RESUBMIT":                RED,
    "FAIL":                    RED,
    "REFER":                   RED,
}


async def run_audit(doc_type: str, text: str) -> str:
    if not model:
        return "AI model not available. Please set GEMINI_API_KEY."
    prompt_template = PROMPT_MAP.get(doc_type, MASTERS_PROMPT)
    prompt = prompt_template.format(text=text[:20000])
    try:
        response = await model.generate_content_async(
            prompt,
            generation_config={"temperature": 0.3, "max_output_tokens": 8192},
            safety_settings=[
                {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            ]
        )
        return response.text.strip()
    except Exception as e:
        print(f"Audit AI error: {e}")
        return f"Audit could not be completed: {e}"


# ═══════════════════════════════════════════════════════════════
#  PDF REPORT BUILDER
# ═══════════════════════════════════════════════════════════════
def build_styles():
    def S(name, **kw):
        kw.setdefault("fontName", "Times-Roman")
        kw.setdefault("fontSize", 11)
        kw.setdefault("leading", 16)
        kw.setdefault("textColor", BLACK)
        kw.setdefault("alignment", TA_JUSTIFY)
        return ParagraphStyle(name, **kw)

    return {
        "Cover":       S("Cover", fontName="Helvetica-Bold", fontSize=22,
                         leading=28, textColor=NAVY, alignment=TA_LEFT),
        "CoverSub":    S("CoverSub", fontName="Helvetica", fontSize=13,
                         leading=20, textColor=ACCENT, alignment=TA_LEFT),
        "Meta":        S("Meta", fontName="Helvetica", fontSize=10,
                         leading=16, textColor=colors.HexColor("#444444"),
                         alignment=TA_LEFT),
        "TypeBadge":   S("TypeBadge", fontName="Helvetica-Bold", fontSize=11,
                         leading=16, textColor=WHITE, alignment=TA_CENTER),
        "SectionHead": S("SectionHead", fontName="Helvetica-Bold", fontSize=13,
                         leading=20, textColor=NAVY, alignment=TA_LEFT,
                         spaceBefore=18, spaceAfter=6),
        "Body":        S("Body", leading=18, spaceAfter=8),
        "BodyBold":    S("BodyBold", fontName="Times-Bold", leading=18, spaceAfter=8),
        "Verdict":     S("Verdict", fontName="Helvetica-Bold", fontSize=14,
                         leading=20, textColor=WHITE, alignment=TA_CENTER),
        "VerdictSub":  S("VerdictSub", fontName="Helvetica", fontSize=10,
                         leading=14, textColor=WHITE, alignment=TA_CENTER),
        "TblHdr":      S("TblHdr", fontName="Helvetica-Bold", fontSize=10,
                         leading=14, textColor=NAVY, alignment=TA_LEFT),
        "TblBody":     S("TblBody", fontName="Times-Roman", fontSize=10,
                         leading=14, textColor=BLACK, alignment=TA_LEFT),
        "Footer":      S("Footer", fontName="Helvetica-Oblique", fontSize=9,
                         leading=13, textColor=colors.HexColor("#888888"),
                         alignment=TA_CENTER),
        "Bullet":      S("Bullet", fontName="Times-Roman", fontSize=11,
                         leading=17, textColor=BLACK, leftIndent=16,
                         alignment=TA_JUSTIFY),
        "Critical":    S("Critical", fontName="Helvetica-Bold", fontSize=11,
                         leading=16, textColor=RED, alignment=TA_LEFT),
    }


def HR(color=RULE, thickness=0.5, before=6, after=8):
    return HRFlowable(width="100%", thickness=thickness,
                      color=color, spaceBefore=before, spaceAfter=after)


def info_box(styles, rows):
    """Key-value box for document metadata."""
    data = [[Paragraph(f'<font name="Helvetica-Bold">{k}</font>', styles["Meta"]),
             Paragraph(v, styles["Meta"])] for k, v in rows]
    t = Table(data, colWidths=[4.5*cm, W - 5*cm - 4.5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), BOX_BG),
        ("BOX",           (0, 0), (-1, -1), 0.5, RULE),
        ("INNERGRID",     (0, 0), (-1, -1), 0.3, RULE),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
    ]))
    return t


def verdict_box(styles, verdict_text, color):
    """Coloured verdict banner."""
    top_text = Paragraph(verdict_text.upper(), styles["Verdict"])
    sub_text  = Paragraph("Examiner's Recommended Outcome", styles["VerdictSub"])
    t = Table([[top_text], [sub_text]],
              colWidths=[W - 5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), color),
        ("TOPPADDING",    (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("LEFTPADDING",   (0, 0), (-1, -1), 20),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 20),
        ("ROUNDEDCORNERS", [6]),
    ]))
    return t


def parse_audit_sections(audit_text: str) -> list[dict]:
    """
    Parse the AI output into sections.
    Looks for patterns like:
      SECTION N — TITLE
      or
      **SECTION N — TITLE**
    Returns list of {title, body}.
    """
    pattern = re.compile(
        r'(?:^|\n)\s*(?:\*{0,2})?SECTION\s+\d+\s*[—–-]+\s*([^\n\*]+)(?:\*{0,2})?',
        re.IGNORECASE
    )
    matches = list(pattern.finditer(audit_text))
    if not matches:
        return [{"title": "Full Audit Report", "body": audit_text.strip()}]

    sections = []
    for i, m in enumerate(matches):
        title = m.group(1).strip().rstrip("*").strip()
        start = m.end()
        end   = matches[i + 1].start() if i + 1 < len(matches) else len(audit_text)
        body  = audit_text[start:end].strip()
        sections.append({"title": title, "body": body})
    return sections


def detect_verdict(audit_text: str) -> tuple[str, colors.Color]:
    """Extract verdict and its colour from the audit text."""
    text_upper = audit_text.upper()
    verdicts = [
        ("READY FOR SUBMISSION",    GREEN),
        ("PASS WITH MINOR CORRECTIONS", AMBER),
        ("PASS WITH MINOR",         AMBER),
        ("MINOR REVISIONS REQUIRED", AMBER),
        ("MAJOR REVISIONS REQUIRED", RED),
        ("MAJOR REVISIONS",         RED),
        ("REFER (RESUBMIT)",        RED),
        ("REJECT AND REWRITE",      RED),
        ("REJECT",                  RED),
        ("RESUBMIT",                RED),
        ("FAIL",                    RED),
        ("PASS",                    GREEN),
    ]
    for label, col in verdicts:
        if label in text_upper:
            return label.title(), col
    return "Review Complete", ACCENT


def body_to_paragraphs(styles, body_text: str) -> list:
    """Convert section body text into flowable paragraphs."""
    items = []
    for line in body_text.split("\n"):
        line = line.strip()
        if not line:
            items.append(Spacer(1, 4))
            continue

        # Strip markdown bold/italic
        line = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', line)
        line = re.sub(r'\*(.+?)\*',     r'<i>\1</i>', line)
        line = re.sub(r'__(.+?)__',     r'<b>\1</b>', line)

        # Bullet points
        if re.match(r'^[-•·]\s+', line):
            text = re.sub(r'^[-•·]\s+', '', line)
            items.append(Paragraph(f"• {text}", styles["Bullet"]))
        # Numbered list
        elif re.match(r'^\d+[\.\)]\s+', line):
            items.append(Paragraph(line, styles["Bullet"]))
        # Bold lead (key: value pattern)
        elif re.match(r'^<b>', line):
            items.append(Paragraph(line, styles["Body"]))
        else:
            items.append(Paragraph(line, styles["Body"]))
    return items


def generate_pdf(
    filename: str,
    audit_id: str,
    doc_type: str,
    classification: dict,
    audit_text: str,
) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2.5*cm, rightMargin=2.5*cm,
        topMargin=2.2*cm,  bottomMargin=2.2*cm,
        title="Academic Audit Report",
        author="PhD Supervisor Audit System",
    )

    S = build_styles()
    story = []

    # ── Page 1: Cover ─────────────────────────────────────────
    story.append(Spacer(1, 0.5*cm))

    # Document type badge (coloured bar)
    type_label = TYPE_LABELS.get(doc_type, doc_type)
    badge_color = {
        "JOURNAL_ARTICLE": colors.HexColor("#0A3D6B"),
        "UNDERGRADUATE":   colors.HexColor("#3A5A1C"),
        "MASTERS":         colors.HexColor("#6B3A00"),
        "PHD":             colors.HexColor("#4A0A0A"),
    }.get(doc_type, NAVY)

    badge_table = Table(
        [[Paragraph(type_label.upper(), S["TypeBadge"])]],
        colWidths=[W - 5*cm]
    )
    badge_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), badge_color),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING",   (0, 0), (-1, -1), 14),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 14),
    ]))
    story.append(badge_table)
    story.append(Spacer(1, 14))

    story.append(Paragraph("ACADEMIC AUDIT REPORT", S["Cover"]))
    story.append(Paragraph("PhD Supervisor Critical Assessment", S["CoverSub"]))
    story.append(Spacer(1, 6))
    story.append(HR(color=NAVY, thickness=1.5, before=0, after=10))

    # Metadata box
    title   = classification.get("title", "UNKNOWN")
    authors = classification.get("authors", "UNKNOWN")
    field   = classification.get("field", "UNKNOWN")
    conf    = classification.get("confidence", "N/A")
    signals = classification.get("signals", [])

    story.append(info_box(S, [
        ("Document",   filename),
        ("Detected As", type_label),
        ("Confidence",  conf),
        ("Title",       title[:90] + ("…" if len(title) > 90 else "")),
        ("Author(s)",   authors[:80] + ("…" if len(authors) > 80 else "")),
        ("Field",       field),
        ("Audit ID",    audit_id),
        ("Date",        datetime.now().strftime("%d %B %Y, %H:%M")),
    ]))

    story.append(Spacer(1, 10))

    # Classification signals
    if signals:
        story.append(Paragraph("Classification Signals:", S["BodyBold"]))
        for sig in signals:
            story.append(Paragraph(f"• {sig}", S["Bullet"]))
    story.append(Spacer(1, 8))

    # Verdict banner
    verdict_label, verdict_color = detect_verdict(audit_text)
    story.append(verdict_box(S, verdict_label, verdict_color))
    story.append(Spacer(1, 10))
    story.append(HR())

    # ── Audit Sections ─────────────────────────────────────────
    sections = parse_audit_sections(audit_text)

    for sec in sections:
        block = []
        block.append(Paragraph(sec["title"].upper(), S["SectionHead"]))
        block.append(HR(color=ACCENT, thickness=0.8, before=0, after=8))
        block.extend(body_to_paragraphs(S, sec["body"]))
        block.append(Spacer(1, 6))
        story.append(KeepTogether(block[:4]))  # keep heading + first para together
        story.extend(block[4:])

    # ── Final footer note ──────────────────────────────────────
    story.append(HR())
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        "This report was generated by the PhD Supervisor Audit System. "
        "It is AI-assisted and intended to support — not replace — human academic judgement. "
        "All findings should be reviewed by a qualified supervisor before communicating to candidates.",
        S["Footer"]
    ))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


def generate_summary_txt(
    filename: str,
    audit_id: str,
    doc_type: str,
    classification: dict,
    audit_text: str,
) -> str:
    """Plain-text summary for quick reading."""
    verdict, _ = detect_verdict(audit_text)
    lines = [
        "=" * 70,
        "  PhD SUPERVISOR AUDIT SYSTEM — PLAIN TEXT SUMMARY",
        "=" * 70,
        f"File:        {filename}",
        f"Audit ID:    {audit_id}",
        f"Date:        {datetime.now().strftime('%d %B %Y %H:%M')}",
        f"Type:        {TYPE_LABELS.get(doc_type, doc_type)}",
        f"Title:       {classification.get('title', 'UNKNOWN')}",
        f"Author(s):   {classification.get('authors', 'UNKNOWN')}",
        f"Field:       {classification.get('field', 'UNKNOWN')}",
        f"Verdict:     {verdict.upper()}",
        "",
        "-" * 70,
        "FULL AUDIT REPORT",
        "-" * 70,
        "",
        audit_text,
        "",
        "-" * 70,
        "END OF REPORT",
        "-" * 70,
    ]
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════
#  MAIN ENDPOINT
# ═══════════════════════════════════════════════════════════════
@app.post("/audit")
async def audit_document(file: UploadFile = File(...)):
    content  = await file.read()
    filename = file.filename or "document"

    audit_hash = hashlib.md5(content).hexdigest()[:10].upper()
    audit_id   = f"SUP-{audit_hash}"

    # 1. Extract text
    text = extract_text(content, filename)
    if not text.strip():
        text = "[Document appears to be image-based or empty — text extraction failed]"

    # 2. Classify
    classification = await classify_document(text)
    doc_type = classification.get("type", "MASTERS")

    # 3. Run type-specific audit
    audit_text = await run_audit(doc_type, text)

    # 4. Generate PDF report
    pdf_bytes = generate_pdf(filename, audit_id, doc_type, classification, audit_text)

    # 5. Generate plain-text summary
    summary_txt = generate_summary_txt(filename, audit_id, doc_type, classification, audit_text)

    # 6. Package as ZIP
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Audit_Report.pdf",  pdf_bytes)
        zf.writestr("Audit_Summary.txt", summary_txt.encode())
    tmp.close()

    return FileResponse(
        path=tmp.name,
        media_type="application/zip",
        filename=f"SupervisorAudit_{audit_id}.zip",
        background=BackgroundTask(lambda: os.unlink(tmp.name)),
        headers={"Access-Control-Expose-Headers": "Content-Disposition"},
    )


@app.get("/health")
async def health():
    return {
        "status": "ok",
        "ai_available": model is not None,
        "version": "2.0.0"
    }
