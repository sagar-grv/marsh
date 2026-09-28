"""
server.py - Enterprise FastAPI Backend for Marsh AI Pitch Generator & Auditor
=============================================================================
Provides REST API endpoints for:
1. Health check & document registry
2. Multi-agent pitch generation, risk mapping, compliance auditing, and export generation
3. Secure file download for PPTX, CSV, JSON, and Markdown compliance assets
4. Static file mounting for the Tailwind/Vanilla JS frontend
"""

import os
import re
import csv
import json
import logging
from io import StringIO
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Dict, Any

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from fpdf import FPDF

# Core multi-agent logic
from core_agents import (
    generateCompanyProfile,
    generateMarketingPitch,
    auditPitchContent,
    export_to_pptx,
    TARGET_PDFS,
    CompanyProfile,
    PitchDeck,
    AuditReport,
)

# --- Logging Setup ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("MarshServer")

# --- Directory Constants ---
BASE_DIR = Path(__file__).parent.resolve()
OUTPUTS_DIR = BASE_DIR / "outputs"
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
FRONTEND_DIR = BASE_DIR / "frontend"

app = FastAPI(
    title="Marsh AI Corporate Pitch Generator API",
    description="Multi-Agent RAG platform for corporate risk placement and policy compliance auditing.",
    version="2.0.0",
)

# Enable CORS for local development and external UI access
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =============================================================================
# Request / Response Schemas
# =============================================================================
class GenerateRequest(BaseModel):
    company_name: str = Field(..., description="Target corporate client name")
    selected_documents: List[str] = Field(
        default_factory=lambda: list(TARGET_PDFS),
        description="List of brochure PDF filenames to ground RAG retrieval",
    )
    llm_provider: str = Field(
        default="Groq (Llama 3.3 70B)",
        description="LLM provider: 'Groq (Llama 3.3 70B)' or 'OpenAI (gpt-4o-mini)'",
    )


# =============================================================================
# Helper Functions: Audit Exporters (CSV, PDF)
# =============================================================================
def generate_audit_csv(audit_report: AuditReport, company_name: str) -> str:
    """Generate CSV string for Excel / PowerBI audit ingestion."""
    buffer = StringIO()
    writer = csv.writer(buffer)
    writer.writerow([
        "Claim_ID",
        "Company_Name",
        "Extracted_Claim",
        "Core_Policy_Feature",
        "Stated_Limit_or_Rule",
        "Client_Application",
        "Auditor_Status",
        "Confidence_Score",
        "Evidence_Snippet",
        "Source_Document",
    ])
    for idx, c in enumerate(audit_report.claims, 1):
        writer.writerow([
            idx,
            company_name,
            c.claim,
            c.core_policy_feature,
            c.stated_limit_or_rule,
            c.client_application,
            c.status,
            f"{c.confidence_score:.2f}",
            c.evidence_snippet.replace("\n", " ").strip(),
            c.source_document,
        ])
    return buffer.getvalue()


def clean_pdf_text(text: str) -> str:
    """
    Prevents fpdf crashes from Indian Rupee symbols, smart quotes, dashes,
    and unhandled Unicode characters.
    """
    if not text:
        return ""
    replacements = {
        "₹": "INR ",
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "–": "-",
        "—": "-",
        "…": "...",
        "\u200b": "",
        "\u2011": "-",
        "\r\n": " ",
        "\n": " ",
        "\t": " ",
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    # Strip any remaining control chars or invalid sequences
    return re.sub(r"\s+", " ", text).strip()


class MarshAuditPDF(FPDF):
    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.font_family_name = "Helvetica"
        self._setup_fonts()

    def _setup_fonts(self):
        """Attempts to register Windows Arial Unicode TrueType fonts, else falls back to Helvetica."""
        arial_regular = "C:/Windows/Fonts/arial.ttf"
        arial_bold = "C:/Windows/Fonts/arialbd.ttf"
        arial_italic = "C:/Windows/Fonts/ariali.ttf"

        if os.path.exists(arial_regular) and os.path.exists(arial_bold):
            try:
                self.add_font("ArialUnicode", "", arial_regular)
                self.add_font("ArialUnicode", "B", arial_bold)
                if os.path.exists(arial_italic):
                    self.add_font("ArialUnicode", "I", arial_italic)
                self.font_family_name = "ArialUnicode"
            except Exception as e:
                logger.warning(f"Could not load system Arial font: {e}. Falling back to standard Helvetica.")
                self.font_family_name = "Helvetica"
        else:
            self.font_family_name = "Helvetica"

    def header(self):
        # Marsh Navy Header Band on every page: RGB(0, 32, 91)
        self.set_fill_color(0, 32, 91)
        self.rect(0, 0, 210, 24, "F")
        self.set_text_color(255, 255, 255)

        # White title
        self.set_font(self.font_family_name, "B", 14)
        self.set_xy(15, 5)
        self.cell(0, 7, "MARSH RISK ADVISORY", 0, 1, "L")

        # Subtitle
        self.set_font(self.font_family_name, "", 9)
        self.set_xy(15, 12)
        self.cell(0, 6, "AI Pitch Compliance Audit Report", 0, 1, "L")
        self.ln(12)

    def footer(self):
        self.set_y(-15)
        self.set_font(self.font_family_name, "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}} | CONFIDENTIAL - MARSH INTERNAL USE ONLY", 0, 0, "C")

    def safe_text(self, text: str) -> str:
        cleaned = clean_pdf_text(str(text) if text is not None else "")
        if self.font_family_name == "Helvetica":
            return cleaned.encode("latin-1", "ignore").decode("latin-1")
        return cleaned

    def section_heading(self, label: str):
        self.set_font(self.font_family_name, "B", 12)
        self.set_text_color(0, 32, 91)
        self.cell(0, 8, self.safe_text(label), 0, 1, "L")
        self.set_draw_color(0, 163, 224)
        self.set_line_width(0.4)
        self.line(15, self.get_y(), 195, self.get_y())
        self.ln(4)


def generate_audit_pdf(audit_report, company_name: str, baseline_docs: Optional[List[str]] = None, output_path: str = "") -> str:
    """
    Renders an executive 4-part compliance audit memo in compliance with Marsh standards:
    1. Executive Summary
    2. Claim Traceability Matrix
    3. Detailed Claim Evidence
    4. Advisor Decision Framework
    """
    if not output_path and baseline_docs and isinstance(baseline_docs, str):
        output_path = baseline_docs
        baseline_docs = []

    pdf = MarshAuditPDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.set_margins(left=15, top=26, right=15)
    pdf.add_page()

    fn = pdf.font_family_name
    claims = getattr(audit_report, "claims", []) or []
    score = getattr(audit_report, "deck_confidence_score", 0.0)
    summary = getattr(audit_report, "summary", "No summary provided.")

    # Determine Compliance Status
    if score >= 0.85:
        status = "PASS"
    elif score >= 0.70:
        status = "REVIEW REQUIRED"
    else:
        status = "FAIL"

    # =========================================================================
    # Section 1: Executive Summary
    # =========================================================================
    pdf.section_heading("Section 1: Executive Summary")

    info_data = [
        ("Client Company:", company_name),
        ("Generated On:", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Policy Documents Used:", ", ".join(baseline_docs) if baseline_docs else "All 4 Ingested Policy Brochures"),
        ("Overall Confidence Score:", f"{int(round(score * 100))}%"),
        ("Compliance Status:", status),
    ]

    for label, val in info_data:
        pdf.set_font(fn, "B", 10)
        pdf.set_text_color(0, 32, 91)
        pdf.cell(55, 6, label, 0, 0)
        
        pdf.set_font(fn, "B" if label == "Compliance Status:" else "", 10)
        if label == "Compliance Status:":
            if status == "PASS":
                pdf.set_text_color(16, 149, 193)  # Marsh Cyan
            elif status == "REVIEW REQUIRED":
                pdf.set_text_color(217, 119, 6)   # Amber
            else:
                pdf.set_text_color(220, 38, 38)   # Red
        else:
            pdf.set_text_color(50, 50, 50)
            
        pdf.cell(0, 6, pdf.safe_text(val), 0, 1)

    pdf.ln(3)
    pdf.set_font(fn, "B", 10)
    pdf.set_text_color(0, 32, 91)
    pdf.cell(0, 6, "Auditor Summary:", 0, 1)

    pdf.set_font(fn, "", 9.5)
    pdf.set_text_color(50, 50, 50)
    pdf.multi_cell(0, 5, pdf.safe_text(summary))
    pdf.ln(6)

    # =========================================================================
    # Section 2: Claim Traceability Matrix
    # =========================================================================
    pdf.section_heading("Section 2: Claim Traceability Matrix")

    # Table layout with auto-wrapping and pagination (Page width: 210, margins: 15+15=30, printable=180)
    col_widths = (22, 58, 48, 32, 20)
    headers = ["Status", "Claim", "Core Policy Feature", "Source Document", "Confidence"]

    try:
        with pdf.table(col_widths=col_widths) as table:
            # Header Row
            h_row = table.row()
            pdf.set_font(fn, "B", 8.5)
            for title in headers:
                h_row.cell(title)

            # Data Rows
            pdf.set_font(fn, "", 8)
            for c in claims:
                row = table.row()
                raw_st = str(getattr(c, "status", "Unknown"))
                is_ver = "Verified" in raw_st or raw_st.upper() == "PASS"
                st_label = "VERIFIED" if is_ver else "FLAGGED"

                c_text = getattr(c, "claim", "")
                c_feat = getattr(c, "core_policy_feature", "N/A")
                c_src = getattr(c, "source_document", "").replace(".pdf", "")
                c_conf = f"{int(round(getattr(c, 'confidence_score', 0.0) * 100))}%"

                row.cell(pdf.safe_text(st_label))
                row.cell(pdf.safe_text(c_text))
                row.cell(pdf.safe_text(c_feat))
                row.cell(pdf.safe_text(c_src))
                row.cell(pdf.safe_text(c_conf))
    except Exception as e:
        logger.warning(f"Table rendering fallback: {e}")
        pdf.set_font(fn, "", 9)
        pdf.set_text_color(50, 50, 50)
        for idx, c in enumerate(claims, 1):
            st = "VERIFIED" if "Verified" in str(getattr(c, "status", "")) else "FLAGGED"
            pdf.multi_cell(0, 5, pdf.safe_text(f"[{st}] Claim {idx}: {getattr(c, 'claim', '')} | Feature: {getattr(c, 'core_policy_feature', '')}"))

    pdf.ln(6)

    # =========================================================================
    # Section 3: Detailed Claim Evidence
    # =========================================================================
    pdf.section_heading("Section 3: Detailed Claim Evidence")

    for i, claim in enumerate(claims, 1):
        raw_status = str(getattr(claim, "status", "Unknown"))
        is_verified = "Verified" in raw_status or raw_status.upper() == "PASS"

        # Claim header & Status badge text
        pdf.set_font(fn, "B", 10.5)
        pdf.set_text_color(0, 32, 91)
        pdf.cell(30, 6, f"Claim ID #{i}:", 0, 0)

        if is_verified:
            pdf.set_text_color(16, 149, 193)  # Cyan
            badge_text = "VERIFIED"
        else:
            pdf.set_text_color(220, 38, 38)   # Red
            badge_text = "FLAGGED"

        pdf.cell(0, 6, f"[{badge_text}]", 0, 1)

        # Field Helper
        def render_field(title: str, content: str, is_italic: bool = False):
            pdf.set_font(fn, "B", 9)
            pdf.set_text_color(0, 32, 91)
            pdf.cell(48, 5, title, 0, 0)
            pdf.set_font(fn, "I" if is_italic else "", 9)
            pdf.set_text_color(80, 80, 80) if is_italic else pdf.set_text_color(50, 50, 50)
            pdf.multi_cell(0, 5, pdf.safe_text(content))
            pdf.ln(1)

        render_field("Extracted Claim:", getattr(claim, "claim", ""))
        render_field("Core Policy Feature:", getattr(claim, "core_policy_feature", "N/A"))
        render_field("Stated Limit:", getattr(claim, "stated_limit_or_rule", "Standard policy terms"))
        render_field("Client Application:", getattr(claim, "client_application", "N/A"))
        
        ev_snip = getattr(claim, "evidence_snippet", "No direct evidence found.")
        render_field("Evidence Snippet:", f'"{ev_snip.strip()}"', is_italic=True)
        render_field("Source Document:", getattr(claim, "source_document", "N/A"))

        # Auditor Note
        auditor_note = getattr(claim, "auditor_note", "") or (
            "Auditor confirmed exact clause match in knowledge base." if is_verified
            else "Potential hallucination or policy term mismatch. Review required."
        )
        render_field("Auditor Note:", auditor_note)
        pdf.ln(3)

    # =========================================================================
    # Section 4: Advisor Decision Framework
    # =========================================================================
    # Ensure fresh page or clean break for Decision Framework
    if pdf.get_y() > 210:
        pdf.add_page()
    else:
        pdf.ln(4)

    pdf.section_heading("Section 4: Advisor Decision Framework")

    framework_rules = [
        ("APPROVE", "If Overall Confidence >= 85% and no critical hallucinations are detected. The pitch deck is verified against policy documents and cleared for client distribution."),
        ("EDIT", "If numeric limits or benefit names need correction using evidence snippets. Advisor should manually adjust numbers/sub-limits in the generated PowerPoint before distribution."),
        ("REJECT", "If core policy benefit is not present in selected documents or benefits are attributed to an insurer that does not offer them. Regenerate with constrained document selection.")
    ]

    for action, explanation in framework_rules:
        pdf.set_font(fn, "B", 10)
        if action == "APPROVE":
            pdf.set_text_color(16, 149, 193)
        elif action == "EDIT":
            pdf.set_text_color(217, 119, 6)
        else:
            pdf.set_text_color(220, 38, 38)

        pdf.cell(28, 6, action + ":", 0, 0)
        pdf.set_font(fn, "", 9.5)
        pdf.set_text_color(50, 50, 50)
        pdf.multi_cell(0, 5, pdf.safe_text(explanation))
        pdf.ln(2)

    pdf.output(output_path)
    logger.info(f"Generated Marsh compliance audit memo at: '{output_path}'")
    return output_path


# =============================================================================
# API Endpoints
# =============================================================================
@app.get("/api/health")
def health_check():
    """Liveness probe."""
    return {
        "status": "ok",
        "timestamp": datetime.now().isoformat(),
        "service": "Marsh AI Pitch Generator API",
    }


@app.get("/api/documents")
def get_documents():
    """Lists available baseline policy documents in the system."""
    docs = []
    for pdf_name in TARGET_PDFS:
        in_root = (BASE_DIR / pdf_name).is_file()
        in_folder = (BASE_DIR / "Policy Documents" / pdf_name).is_file()
        exists = in_root or in_folder
        docs.append({
            "filename": pdf_name,
            "display_name": pdf_name.replace(".pdf", ""),
            "available": exists,
        })
    return {"documents": docs}


@app.post("/api/generate")
def generate_pitch(req: GenerateRequest):
    """
    Executes the multi-agent orchestration:
    1. Company Profiler (Web search / assumption)
    2. Query Reformulator + RAG Retriever Tool (BGE normalized)
    3. Pitch Generator Agent (4-slide structured deck)
    4. Compliance Auditor Agent (Fact verification)
    5. Exporters (PPTX, CSV, PDF, JSON)
    """
    company_name = req.company_name.strip()
    if not company_name:
        raise HTTPException(status_code=400, detail="Company name cannot be empty.")

    selected_docs = [d for d in req.selected_documents if d.strip()]
    if not selected_docs:
        raise HTTPException(status_code=400, detail="At least one policy brochure must be selected.")

    logger.info(f"Incoming pitch request for: '{company_name}' with {len(selected_docs)} brochures.")

    try:
        # 1. Company Profiler
        logger.info(f"Step 1: Profiling company '{company_name}'...")
        profile = generateCompanyProfile(company_name, llm_provider=req.llm_provider)

        # 2. Pitch Generator + RAG Retrieval
        logger.info(f"Step 2: Generating pitch deck & querying policy vector store...")
        pitch_deck, retrieved_chunks = generateMarketingPitch(
            profile=profile,
            selected_documents=selected_docs,
            llm_provider=req.llm_provider,
        )

        # 3. Compliance Auditor Agent
        logger.info(f"Step 3: Auditing pitch claims against policy knowledge base...")
        audit_report = auditPitchContent(
            pitch_deck=pitch_deck,
            selected_documents=selected_docs,
            llm_provider=req.llm_provider,
        )

        # 4. Generate Export Files in outputs/
        safe_slug = re.sub(r"[^\w\s-]", "", company_name).strip().replace(" ", "_")
        if not safe_slug:
            safe_slug = "Corporate_Client"

        pptx_filename = f"{safe_slug}_Pitch_Deck.pptx"
        csv_filename = f"Audit_Report_{safe_slug}.csv"
        json_filename = f"Audit_Report_{safe_slug}.json"
        pdf_filename = f"{safe_slug}_Audit_Report.pdf"

        pptx_path = OUTPUTS_DIR / pptx_filename
        csv_path = OUTPUTS_DIR / csv_filename
        json_path = OUTPUTS_DIR / json_filename
        pdf_path = OUTPUTS_DIR / pdf_filename

        # Render 16:9 McKinsey/Marsh PPTX
        export_to_pptx(pitch_deck, output_path=str(pptx_path), profile=profile)

        # Render CSV
        csv_content = generate_audit_csv(audit_report, company_name)
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(csv_content)

        # Render JSON
        json_payload = {
            "client_organization": company_name,
            "evaluation_date": datetime.now().isoformat(),
            "overall_confidence_score": audit_report.deck_confidence_score,
            "summary": audit_report.summary,
            "baseline_documents_grounded": selected_docs,
            "claims": [c.model_dump() for c in audit_report.claims],
        }
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_payload, f, indent=2)

        # Render Markdown Audit Report
        md_filename = f"Audit_Report_{safe_slug}.md"
        md_path = OUTPUTS_DIR / md_filename
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(f"# Marsh Risk Advisory - Compliance & Hallucination Audit Report\n\n")
            f.write(f"**Client Organization:** {company_name}  \n")
            f.write(f"**Evaluation Date:** {datetime.now().strftime('%B %d, %Y')}  \n")
            f.write(f"**Overall Deck Confidence Score:** {audit_report.deck_confidence_score * 100:.1f}%  \n")
            f.write(f"**Baseline Documents Grounded:** {', '.join(selected_docs)}  \n\n---\n\n")
            f.write(f"## 1. Executive Compliance Summary\n{audit_report.summary}\n\n")
            verdict = "PASS - COMPLIANT" if audit_report.deck_confidence_score >= 0.70 else "REVIEW REQUIRED"
            f.write(f"**Compliance Verdict:** {verdict}\n\n---\n\n")
            f.write(f"## 2. Claim-by-Claim Verification Breakdown\n\n")
            for i, c in enumerate(audit_report.claims, 1):
                status_badge = "[PASS] VERIFIED" if "Verified" in c.status else "[FLAGGED] HALLUCINATION/UNVERIFIED"
                f.write(f"### Claim {i}: {status_badge} (Confidence: {c.confidence_score:.2f})\n")
                f.write(f"- **Original Pitch Claim:** \"{c.claim}\"\n")
                f.write(f"- **Core Policy Feature:** `{c.core_policy_feature}`\n")
                f.write(f"- **Stated Limit or Rule:** `{c.stated_limit_or_rule}`\n")
                f.write(f"- **Client Application / Risk Mapping:** `{c.client_application}`\n")
                f.write(f"- **Policy Brochure Citation:** `{c.source_document}`\n")
                f.write(f"- **Auditor Evidence & Rationale:**\n  > {c.evidence_snippet.strip()}\n\n")

        # Render Corporate PDF Audit Report
        pdf_full_path = generate_audit_pdf(audit_report, company_name, selected_docs, str(pdf_path))

        logger.info(f"Generated all assets for '{company_name}' in '{OUTPUTS_DIR}'.")

        # Response payload
        return {
            "success": True,
            "company_name": company_name,
            "profile": profile.model_dump(),
            "pitch_deck": pitch_deck.model_dump(),
            "audit_report": audit_report.model_dump(),
            "retrieved_chunks": retrieved_chunks,
            "pdf_file_path": str(pdf_full_path),
            "downloads": {
                "pptx": {
                    "filename": pptx_filename,
                    "url": f"/api/download/pptx/{pptx_filename}",
                    "file_path": str(pptx_path),
                },
                "pdf": {
                    "filename": pdf_filename,
                    "url": f"/api/download/pdf/{pdf_filename}",
                    "file_path": str(pdf_full_path),
                },
                "csv": {
                    "filename": csv_filename,
                    "url": f"/api/download/csv/{csv_filename}",
                    "file_path": str(csv_path),
                },
                "json": {
                    "filename": json_filename,
                    "url": f"/api/download/json/{json_filename}",
                    "file_path": str(json_path),
                },
            },
        }

    except Exception as e:
        logger.exception(f"Error processing pitch generation: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/download/{file_type}/{filename}")
def download_asset(file_type: str, filename: str):
    """
    Secure file download endpoint preventing directory traversal attacks.
    """
    # Prevent directory traversal
    clean_name = os.path.basename(filename)
    target_path = (OUTPUTS_DIR / clean_name).resolve()

    if not target_path.is_file() or not str(target_path).startswith(str(OUTPUTS_DIR.resolve())):
        raise HTTPException(status_code=404, detail="Requested file not found or access denied.")

    media_types = {
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "pdf": "application/pdf",
        "csv": "text/csv",
        "json": "application/json",
    }
    media_type = media_types.get(file_type.lower(), "application/octet-stream")

    return FileResponse(
        path=target_path,
        media_type=media_type,
        filename=clean_name,
    )


# =============================================================================
# Mount Static Frontend
# =============================================================================
if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="127.0.0.1", port=8000, reload=True)
