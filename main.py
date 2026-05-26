"""
ThesisSifu Pro — Multi-Agent Thesis Panel
==========================================
Three-output audit system for academic documents.

Agent 1 (Gemini Flash Lite):
  - Classifies document type
  - Processes each chapter paragraph-by-paragraph
  - Produces inline comments: severity + literature type + theory needed

Agent 2 (Claude Sonnet 4.6):
  - Synthesises all chapter summaries
  - Holistic examiner report: chapter alignment, golden thread,
    Scopus tone, citation integrity, viva questions, verdict

Outputs (ZIP):
  1_Examiner_Audit_Report.pdf   — holistic examiner critique
  2_Annotated_Thesis.docx       — original doc with inline Word comments
  3_Commentary_Report.pdf       — paragraph-level log with page refs

Endpoint:  POST /audit   multipart/form-data { file: <pdf|docx> }
Returns:   application/zip
"""

from __future__ import annotations

import io, os, re, json, asyncio, zipfile, hashlib, tempfile
from datetime import datetime
from dataclasses import dataclass, field as dc_field
from typing import Optional

import pypdf
from docx import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

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
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False


# ── App ────────────────────────────────────────────────────────
app = FastAPI(title="ThesisSifu Pro — Multi-Agent Thesis Panel", version="3.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"],
                   allow_credentials=False, allow_methods=["*"], allow_headers=["*"])


# ── AI Clients ─────────────────────────────────────────────────
gemini_classifier = None
claude_client     = None

if GEMINI_AVAILABLE:
    gkey = os.environ.get("GEMINI_API_KEY")
    if gkey:
        genai.configure(api_key=gkey)
        gemini_classifier = genai.GenerativeModel("gemini-2.5-flash-lite-preview-06-17")
        print("Gemini Flash Lite ready")

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


# ── Data classes ───────────────────────────────────────────────
@dataclass
class ParagraphComment:
    chapter:          str
    para_index:       int
    page_estimate:    int
    para_excerpt:     str
    severity:         str
    issue:            str
    recommendation:   str
    literature_needed: str
    theory_needed:    str

@dataclass
class ChapterSummary:
    chapter_num:   str
    chapter_title: str
    comments:      list = dc_field(default_factory=list)


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


# ── Chapter splitter ───────────────────────────────────────────
_CH_PAT = re.compile(
    r'(?:^|\n)(?:CHAPTER\s+(\d+|[IVX]+)|(\d+)\.0)\b[:\s\-—–]*([^\n]{0,80})',
    re.IGNORECASE,
)

def split_chapters(text: str) -> list[dict]:
    matches = list(_CH_PAT.finditer(text))
    if not matches:
        size = 6000
        return [{"num": str(i+1), "title": f"Section {i+1}", "text": text[s:s+size]}
                for i, s in enumerate(range(0, len(text), size))]
    chapters = []
    for i, m in enumerate(matches):
        num   = (m.group(1) or m.group(2) or str(i+1)).strip()
        title = (m.group(3) or "").strip()
        start = m.start()
        end   = matches[i+1].start() if i+1 < len(matches) else len(text)
        chapters.append({"num": num, "title": title or f"Chapter {num}",
                         "text": text[start:end].strip()})
    return chapters or [{"num": "1", "title": "Full Document", "text": text}]


# ── Agent 1: Classification ────────────────────────────────────
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
    if not gemini_classifier or not text.strip():
        return default
    try:
        r = await gemini_classifier.generate_content_async(
            _CLS_PROMPT.format(text=text[:5000]))
        raw = re.sub(r"```(?:json)?","", r.text.strip()).strip("`").strip()
        return json.loads(raw)
    except Exception as e:
        print(f"Classification error: {e}")
        return default


# ── Agent 1: Paragraph commentary ─────────────────────────────
_PARA_PROMPT = """
You are a PhD supervisor reviewing Chapter {ch_num} of a {doc_type} titled "{title}".

Identify 5-10 paragraphs needing attention. For each return a JSON object:
{{"para_excerpt":"first 100 chars","severity":"CRITICAL|MODERATE|SUGGESTION",
  "issue":"specific problem","recommendation":"specific fix",
  "literature_needed":"type of research/studies needed — do NOT invent titles",
  "theory_needed":"framework that would help, or none"}}

Return ONLY a JSON array. No markdown fences.

Chapter {ch_num} — {ch_title}:
\"\"\"{text}\"\"\"
"""

async def audit_chapter(ch_num, ch_title, ch_text, doc_type, title) -> list:
    if not gemini_classifier or not ch_text.strip():
        return []
    prompt = _PARA_PROMPT.format(
        ch_num=ch_num, ch_title=ch_title, doc_type=doc_type,
        title=title, text=ch_text[:8000])
    try:
        r = await gemini_classifier.generate_content_async(prompt)
        raw = re.sub(r"```(?:json)?","", r.text.strip()).strip("`").strip()
        items = json.loads(raw)
        if not isinstance(items, list):
            items = [items]
        comments = []
        for idx, it in enumerate(items):
            excerpt = it.get("para_excerpt","")
            pos = ch_text.find(excerpt[:40])
            pg  = max(1, len(ch_text[:pos].split()) // 350 + 1) if pos > 0 else 1
            comments.append(ParagraphComment(
                chapter=f"Chapter {ch_num}",
                para_index=idx, page_estimate=pg,
                para_excerpt=excerpt[:150],
                severity=it.get("severity","MODERATE").upper(),
                issue=it.get("issue",""),
                recommendation=it.get("recommendation",""),
                literature_needed=it.get("literature_needed",""),
                theory_needed=it.get("theory_needed","none"),
            ))
        return comments
    except Exception as e:
        print(f"Paragraph audit error ch{ch_num}: {e}")
        return []


# ── Agent 2: Holistic examiner (Claude Sonnet) ─────────────────
_EXAMINER_SYSTEM = """
You are a senior academic examiner with 25 years of experience.
Produce rigorous, specific, honest examiner-level critique.
You are NOT a grammar checker — you are an intellectual critic.
Respond in the same language as the document.
"""

_EXAMINER_PROMPTS = {
"PHD": """
External examiner for a PhD viva.
Title: "{title}" | Field: {field} | Institution: {institution}

Chapter issue summary:
{summaries}

Write a full examiner's report with EXACTLY these sections:

SECTION 1 — ORIGINAL CONTRIBUTION TO KNOWLEDGE
SECTION 2 — GOLDEN THREAD ANALYSIS
Does title → objectives → RQs → methodology → findings → discussion → conclusion hold together?
SECTION 3 — CHAPTER ALIGNMENT AUDIT
Does each chapter deliver what the previous chapter promised? Cite specific misalignments.
SECTION 4 — THEORETICAL FRAMEWORK COHERENCE
SECTION 5 — METHODOLOGICAL RIGOUR
SECTION 6 — DATA QUALITY & ANALYTICAL DEPTH
SECTION 7 — SCOPUS-LEVEL LANGUAGE & TONE
Assess writing against Q1/Q2 journal standards. Flag specific passages and what needs rewriting.
SECTION 8 — CITATION & REFERENCE INTEGRITY
List all citation problems found across chapters.
SECTION 9 — CRITICAL VIVA QUESTIONS
List 8 specific questions the candidate must defend.
SECTION 10 — EXAMINER'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS REQUIRED / REFER (RESUBMIT) / FAIL
Formal statement 6-8 sentences. Specific conditions before degree is awarded.
""",
"MASTERS": """
Internal reader for a Master's thesis viva.
Title: "{title}" | Field: {field} | Institution: {institution}

Chapter issue summary:
{summaries}

Write a report with EXACTLY these sections:

SECTION 1 — THESIS ALIGNMENT & GOLDEN THREAD
SECTION 2 — CHAPTER ALIGNMENT AUDIT
SECTION 3 — THEORETICAL FRAMEWORK
SECTION 4 — METHODOLOGICAL RIGOUR
SECTION 5 — DATA ROBUSTNESS & INSTRUMENTATION
SECTION 6 — SCOPUS-LEVEL LANGUAGE & TONE
SECTION 7 — CITATION INTEGRITY
SECTION 8 — CRITICAL VIVA QUESTIONS
List 5 questions.
SECTION 9 — EXAMINER'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS REQUIRED / RESUBMIT
Top 5 issues that must be resolved.
""",
"JOURNAL_ARTICLE": """
Peer reviewer for a Scopus journal.
Title: "{title}" | Field: {field}

Section notes:
{summaries}

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
Top 3 critical revisions.
""",
"UNDERGRADUATE": """
Supervisor reviewing an undergraduate FYP.
Title: "{title}" | Field: {field}

Section notes:
{summaries}

Write a report with EXACTLY these sections:

SECTION 1 — SCOPE & RESEARCH QUESTION
SECTION 2 — LITERATURE REVIEW
SECTION 3 — METHODOLOGY
SECTION 4 — ANALYSIS & FINDINGS
SECTION 5 — CRITICAL THINKING
SECTION 6 — WRITING & ACADEMIC CONVENTIONS
SECTION 7 — TOP ISSUES TO FIX
SECTION 8 — SUPERVISOR'S VERDICT
One of: PASS / PASS WITH MINOR CORRECTIONS / MAJOR REVISIONS / FAIL
""",
}

async def run_examiner(doc_type, title, field, inst, chapter_summaries) -> str:
    if not claude_client:
        return "Examiner synthesis unavailable — ANTHROPIC_API_KEY not set."

    summaries_text = ""
    for cs in chapter_summaries:
        n_c = sum(1 for c in cs.comments if c.severity=="CRITICAL")
        n_m = sum(1 for c in cs.comments if c.severity=="MODERATE")
        summaries_text += (f"\nCh.{cs.chapter_num} — {cs.chapter_title}: "
                           f"{len(cs.comments)} comments "
                           f"(Critical:{n_c} Moderate:{n_m})\n")
        for c in cs.comments:
            if c.severity == "CRITICAL":
                summaries_text += f"  [CRITICAL] {c.issue}\n"

    template = _EXAMINER_PROMPTS.get(doc_type, _EXAMINER_PROMPTS["MASTERS"])
    prompt = template.format(title=title, field=field,
                             institution=inst, summaries=summaries_text[:8000])
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
        "Body":    S("Body",leading=18,spaceAfter=6),
        "Bold":    S("Bold",fontName="Times-Bold",leading=18,spaceAfter=6),
        "Bullet":  S("Bullet",leading=17,leftIndent=16,spaceAfter=4),
        "Verdict": S("Verdict",fontName="Helvetica-Bold",fontSize=15,leading=22,textColor=WHITE,alignment=TA_CENTER),
        "VrdSub":  S("VrdSub",fontName="Helvetica",fontSize=10,leading=14,textColor=WHITE,alignment=TA_CENTER),
        "TblH":    S("TblH",fontName="Helvetica-Bold",fontSize=9,leading=13,textColor=NAVY,alignment=TA_LEFT),
        "TblC":    S("TblC",fontName="Times-Roman",fontSize=9,leading=13,textColor=BLACK,alignment=TA_LEFT),
        "Footer":  S("Footer",fontName="Helvetica-Oblique",fontSize=8,leading=12,textColor=colors.HexColor("#999"),alignment=TA_CENTER),
        "Excerpt": S("Excerpt",fontName="Times-Italic",fontSize=9,leading=13,textColor=colors.HexColor("#555"),leftIndent=8),
        "Detail":  S("Detail",fontName="Times-Roman",fontSize=10,leading=15,textColor=BLACK,leftIndent=4),
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

def vbanner(S, text, color):
    t=Table([[Paragraph(text.upper(),S["Verdict"])],[Paragraph("Examiner's Recommended Outcome",S["VrdSub"])]],
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


# ── Output 1: Examiner Report PDF ──────────────────────────────
def build_examiner_pdf(filename, audit_id, doc_type, clf, examiner_text, chs) -> bytes:
    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,leftMargin=2.5*cm,rightMargin=2.5*cm,
                          topMargin=2.2*cm,bottomMargin=2.2*cm)
    S=pdf_styles(); story=[]

    type_label=TYPE_LABELS.get(doc_type,doc_type)
    story.append(type_badge_table(S,type_label.upper(),TYPE_COLORS.get(doc_type,NAVY)))
    story.append(Spacer(1,14))
    story.append(Paragraph("EXAMINER'S AUDIT REPORT",S["Title"]))
    story.append(Paragraph("ThesisSifu Pro — Multi-Agent Thesis Panel",S["Sub"]))
    story.append(HR(c=NAVY,t=1.5,b=4,a=10))

    ti=clf.get("title","UNKNOWN"); au=clf.get("authors","UNKNOWN")
    fi=clf.get("field","UNKNOWN"); ins=clf.get("institution","UNKNOWN")

    story.append(meta_box(S,[
        ("Document",filename),("Detected As",type_label),
        ("Confidence",clf.get("confidence","N/A")),
        ("Title",(ti[:85]+"…") if len(ti)>85 else ti),
        ("Author(s)",(au[:80]+"…") if len(au)>80 else au),
        ("Field",fi),("Institution",ins),
        ("Audit ID",audit_id),("Date",datetime.now().strftime("%d %B %Y, %H:%M")),
        ("Report","1 of 3 — Examiner Audit Report"),
    ]))
    story.append(Spacer(1,10))

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

    vl,vc=detect_verdict(examiner_text)
    story.append(vbanner(S,vl,vc)); story.append(Spacer(1,10)); story.append(HR())

    for sec in parse_secs(examiner_text):
        bl=[Paragraph(sec["title"].upper(),S["SecHead"]),
            HR(c=ACCENT,t=.7,b=0,a=6)]
        bl.extend(body2story(S,sec["body"])); bl.append(Spacer(1,6))
        story.append(KeepTogether(bl[:3])); story.extend(bl[3:])

    # Chapter summary table
    story.append(PageBreak())
    story.append(Paragraph("CHAPTER-BY-CHAPTER ISSUE SUMMARY",S["SecHead"]))
    story.append(HR(c=NAVY,t=1,b=0,a=8))
    td=[[Paragraph(h,S["TblH"]) for h in ["Chapter","Critical","Moderate","Suggestions","Total"]]]
    for cs in chs:
        nc2=sum(1 for c in cs.comments if c.severity=="CRITICAL")
        nm2=sum(1 for c in cs.comments if c.severity=="MODERATE")
        ns2=len(cs.comments)-nc2-nm2
        td.append([Paragraph(f"Ch.{cs.chapter_num} — {cs.chapter_title[:35]}",S["TblC"]),
                   Paragraph(str(nc2),S["TblC"]),Paragraph(str(nm2),S["TblC"]),
                   Paragraph(str(ns2),S["TblC"]),Paragraph(str(nc2+nm2+ns2),S["TblC"])])
    cw=W-5*cm
    t=Table(td,colWidths=[cw*.45,cw*.14,cw*.14,cw*.14,cw*.13])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,BOX_BG]),("BOX",(0,0),(-1,-1),.5,RULE),
        ("INNERGRID",(0,0),(-1,-1),.3,RULE),("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7)]))
    story.append(t)
    story.append(Spacer(1,20)); story.append(HR())
    story.append(Paragraph("ThesisSifu Pro — AI-assisted academic audit. "
        "See Report 2 (Annotated Thesis) and Report 3 (Commentary Log) in this ZIP.",S["Footer"]))

    doc.build(story); buf.seek(0); return buf.getvalue()


# ── Output 2: Annotated DOCX ───────────────────────────────────
def _insert_comment(doc, para, cid, author, text, date_str):
    """Insert a Word comment on a paragraph using lxml."""
    W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # Get or create comments part
    try:
        cp = doc.part.comments
        root = etree.fromstring(cp._blob)
    except Exception:
        xml_str = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        )
        # Attach via relationship
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        uri = PackURI('/word/comments.xml')
        ct  = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
        cp  = Part(uri, ct, xml_str.encode(), doc.part.package)
        doc.part.relate_to(
            cp, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')
        root = etree.fromstring(xml_str.encode())

    # Add <w:comment> element
    comment_el = etree.SubElement(root, f'{{{W_NS}}}comment')
    comment_el.set(f'{{{W_NS}}}id',      str(cid))
    comment_el.set(f'{{{W_NS}}}author',  author)
    comment_el.set(f'{{{W_NS}}}date',    date_str)
    comment_el.set(f'{{{W_NS}}}initials','TS')
    p_el = etree.SubElement(comment_el, f'{{{W_NS}}}p')
    r_el = etree.SubElement(p_el,       f'{{{W_NS}}}r')
    t_el = etree.SubElement(r_el,       f'{{{W_NS}}}t')
    t_el.text = text
    t_el.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    cp._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # Add markers to paragraph XML
    px = para._p
    crs = OxmlElement('w:commentRangeStart'); crs.set(qn('w:id'),str(cid)); px.insert(0,crs)
    cre = OxmlElement('w:commentRangeEnd');   cre.set(qn('w:id'),str(cid)); px.append(cre)
    rr  = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr'); rs = OxmlElement('w:rStyle')
    rs.set(qn('w:val'),'CommentReference'); rpr.append(rs); rr.append(rpr)
    cr  = OxmlElement('w:commentReference'); cr.set(qn('w:id'),str(cid)); rr.append(cr)
    px.append(rr)


def build_annotated_docx(content, filename, audit_id, chs, clf) -> bytes:
    is_docx = filename.lower().endswith(".docx")

    if is_docx:
        doc = DocxDocument(io.BytesIO(content))
    else:
        # PDF: build a fresh DOCX reproducing structure
        doc = DocxDocument()
        doc.add_heading("ThesisSifu Pro — Annotated Commentary", 0)
        doc.add_paragraph(
            f"Source file: {filename}\n"
            "Inline comments are inserted below each flagged paragraph."
        )
        doc.add_page_break()
        for cs in chs:
            doc.add_heading(f"Chapter {cs.chapter_num} — {cs.chapter_title}", 1)
            doc.add_paragraph("[Original text — supervisor comments attached below]")

    date_str = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    cid = 0

    # Build excerpt → comment lookup
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
                sev = {"CRITICAL":"[CRITICAL]","MODERATE":"[MODERATE]",
                       "SUGGESTION":"[SUGGESTION]"}.get(c.severity,"[NOTE]")
                txt = (f"{sev} ThesisSifu Pro\n\n"
                       f"ISSUE: {c.issue}\n\n"
                       f"RECOMMENDATION: {c.recommendation}\n\n"
                       f"LITERATURE NEEDED: {c.literature_needed}\n\n"
                       f"THEORY/FRAMEWORK: {c.theory_needed}")
                try:
                    _insert_comment(doc, para, cid, "ThesisSifu Pro", txt, date_str)
                    cid += 1
                except Exception as e:
                    print(f"Comment insert error {cid}: {e}")
                    cid += 1

    # Header
    hdr = doc.sections[0].header
    if hdr.paragraphs:
        hdr.paragraphs[0].text = (
            f"ThesisSifu Pro | Annotated Thesis | Audit ID: {audit_id} | "
            f"{datetime.now().strftime('%d %B %Y')}"
        )

    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf.getvalue()


# ── Output 3: Commentary Report PDF ───────────────────────────
def build_commentary_pdf(filename, audit_id, doc_type, clf, chs) -> bytes:
    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,leftMargin=2.5*cm,rightMargin=2.5*cm,
                          topMargin=2.2*cm,bottomMargin=2.2*cm)
    S=pdf_styles(); story=[]

    story.append(type_badge_table(S,"PARAGRAPH-LEVEL COMMENTARY REPORT",
                                  TYPE_COLORS.get(doc_type,NAVY)))
    story.append(Spacer(1,14))
    story.append(Paragraph("COMMENTARY REPORT",S["Title"]))
    story.append(Paragraph("ThesisSifu Pro — Paragraph-Level Supervisor Notes",S["Sub"]))
    story.append(HR(c=NAVY,t=1.5,b=4,a=10))

    ti=clf.get("title","UNKNOWN"); au=clf.get("authors","UNKNOWN")
    story.append(meta_box(S,[
        ("Document",filename),("Type",TYPE_LABELS.get(doc_type,doc_type)),
        ("Title",(ti[:85]+"…") if len(ti)>85 else ti),
        ("Author(s)",(au[:80]+"…") if len(au)>80 else au),
        ("Audit ID",audit_id),("Date",datetime.now().strftime("%d %B %Y, %H:%M")),
        ("Report","3 of 3 — Paragraph Commentary Log"),
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

        for i,c in enumerate(cs.comments,1):
            sc = SEV_COLORS.get(c.severity, ACCENT)

            hdr_row = Table(
                [[Paragraph(c.severity,S["Badge"]),Spacer(6,1),
                  Paragraph(f"Comment {i} | {cs.chapter_title[:28]} | Page ~{c.page_estimate}",S["Meta"])]],
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

            detail_box = Table(
                [[Paragraph(
                    f'<b>Issue:</b> {c.issue}<br/><br/>'
                    f'<b>Recommendation:</b> {c.recommendation}<br/><br/>'
                    f'<b>Literature needed:</b> {c.literature_needed}<br/><br/>'
                    f'<b>Theory / Framework:</b> {c.theory_needed}',
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
        "ThesisSifu Pro — Paragraph commentary log. "
        "See Report 1 (Examiner Audit) for holistic verdict and "
        "Report 2 (Annotated Thesis) for inline Word comments.",S["Footer"]))

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

    # 2. Classify
    clf      = await classify_document(full_text)
    doc_type = clf.get("type","MASTERS")
    title    = clf.get("title","UNKNOWN")
    field    = clf.get("field","UNKNOWN")
    inst     = clf.get("institution","UNKNOWN")

    # 3. Split chapters
    chapters = split_chapters(full_text)

    # 4. Agent 1: paragraph audit per chapter (parallel, max 4 at once)
    sem = asyncio.Semaphore(4)
    async def bounded(ch):
        async with sem:
            return await audit_chapter(ch["num"],ch["title"],ch["text"],doc_type,title)

    results = await asyncio.gather(*[bounded(ch) for ch in chapters])

    chs = [ChapterSummary(ch["num"], ch["title"], cmts)
           for ch, cmts in zip(chapters, results)]

    # 5. Agent 2: holistic examiner synthesis
    examiner_text = await run_examiner(doc_type, title, field, inst, chs)

    # 6. Build three outputs
    pdf1  = build_examiner_pdf(filename,  audit_id, doc_type, clf, examiner_text, chs)
    docx2 = build_annotated_docx(content, filename, audit_id, chs, clf)
    pdf3  = build_commentary_pdf(filename, audit_id, doc_type, clf, chs)

    # 7. ZIP and return
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    with zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("1_Examiner_Audit_Report.pdf", pdf1)
        zf.writestr("2_Annotated_Thesis.docx",     docx2)
        zf.writestr("3_Commentary_Report.pdf",      pdf3)
    tmp.close()

    return FileResponse(
        path=tmp.name, media_type="application/zip",
        filename=f"ThesisSifu_Audit_{audit_id}.zip",
        background=BackgroundTask(lambda: os.unlink(tmp.name)),
        headers={"Access-Control-Expose-Headers":"Content-Disposition"},
    )


@app.get("/health")
async def health():
    return {
        "status":                "ok",
        "classifier_available":  gemini_classifier is not None,
        "audit_model_available": claude_client is not None,
        "version":               "3.0.0",
        "outputs":               ["1_Examiner_Audit_Report.pdf",
                                  "2_Annotated_Thesis.docx",
                                  "3_Commentary_Report.pdf"],
    }
