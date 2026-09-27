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


class MarshAuditPDF(FPDF):
    def header(self):
        # Marsh Navy Header
        self.set_fill_color(0, 32, 91)
        self.rect(0, 0, 210, 22, 'F')
        self.set_text_color(255, 255, 255)
        
        self.set_font('Helvetica', 'B', 14)
        self.set_xy(15, 6)
        self.cell(0, 8, 'MARSH RISK ADVISORY', 0, 1, 'L')
        
        self.set_font('Helvetica', '', 9)
        self.set_xy(15, 13)
        self.cell(0, 6, 'AI Pitch Compliance Audit Report', 0, 1, 'L')
        self.ln(12)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}} | CONFIDENTIAL - MARSH INTERNAL USE ONLY', 0, 0, 'C')

    def chapter_title(self, label):
        self.set_font('Helvetica', 'B', 12)
        self.set_text_color(0, 32, 91)
        self.cell(0, 8, label, 0, 1, 'L')
        self.set_draw_color(0, 163, 224)
        self.line(self.get_x(), self.get_y(), self.get_x() + 180, self.get_y())
        self.ln(4)

    def chapter_body(self, text):
        self.set_font('Helvetica', '', 10)
        self.set_text_color(50, 50, 50)
        self.multi_cell(0, 6, text)
        self.ln()


def clean_pdf_text(text: str) -> str:
    """Prevents fpdf crashes from Indian Rupee symbols, smart quotes, and unicode."""
    if not text:
        return ""
    replacements = {
        "₹": "INR ", "’": "'", "‘": "'", "“": '"', "”": '"',
        "–": "-", "—": "-", "…": "...", "\u200b": "", "\u2011": "-", "\n": " "
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text.encode('latin-1', 'ignore').decode('latin-1')


def generate_audit_pdf(audit_report, company_name: str, baseline_docs: Optional[List[str]] = None, output_path: str = ""):
    """
    Renders a 4-part compliance audit memo:
    1. Executive Summary & Confidence Score
    2. Claim Traceability Matrix (fpdf2 table layout)
    3. Detailed Evidence & Auditor Notes (Deep dive)
    4. Advisor Decision Framework (Approve / Edit / Reject)
    """
    # Accommodate flexible argument signature
    if not output_path and baseline_docs and isinstance(baseline_docs, str):
        output_path = baseline_docs
        baseline_docs = []

    pdf = MarshAuditPDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    # --- 1. EXECUTIVE SUMMARY ---
    pdf.chapter_title("1. Executive Summary")
    
    score = getattr(audit_report, 'deck_confidence_score', 0.0)
    summary = getattr(audit_report, 'summary', 'No summary provided.')
    claims = getattr(audit_report, 'claims', [])
    
    score_pct = f"{int(round(score * 100))}%"
    status = "PASS - COMPLIANT" if score >= 0.70 else "REVIEW REQUIRED"
    
    # Info Box
    pdf.set_font('Helvetica', '', 10)
    pdf.set_text_color(50, 50, 50)
    info_data = [
        ["Client Company:", company_name],
        ["Generation Date:", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Overall Confidence Score:", score_pct],
        ["Compliance Status:", status],
        ["Total Claims Audited:", str(len(claims))],
        ["Knowledge Base Grounding:", f"{len(baseline_docs) if baseline_docs else 4} Ingested Brochures"]
    ]
    
    for label, value in info_data:
        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(55, 6, label, 0, 0)
        pdf.set_font('Helvetica', '', 10)
        pdf.cell(0, 6, clean_pdf_text(str(value)), 0, 1)
    pdf.ln(4)
    
    pdf.chapter_body(f"Auditor Summary: {clean_pdf_text(summary)}")

    # --- 2. TRACEABILITY MATRIX ---
    pdf.chapter_title("2. Claim Traceability Matrix")
    
    # Using fpdf2 built-in table for auto-wrapping and pagination
    try:
        with pdf.table(col_widths=(22, 68, 55, 25)) as table:
            # Header Row
            header = table.row()
            for col_name in ["Status", "Claim Extracted", "Source Document", "Confidence"]:
                header.cell(col_name)
            
            # Data Rows
            for claim in claims:
                row = table.row()
                c_status = getattr(claim, 'status', 'Unknown')
                status_txt = "VERIFIED" if "Verified" in str(c_status) else "FLAGGED"
                
                row.cell(clean_pdf_text(status_txt))
                row.cell(clean_pdf_text(getattr(claim, 'claim', '')))
                row.cell(clean_pdf_text(getattr(claim, 'source_document', '').replace(".pdf", "")))
                row.cell(f"{int(round(getattr(claim, 'confidence_score', 0) * 100))}%")
    except Exception as e:
        pdf.chapter_body(f"Table rendering note: {e}")
        
    pdf.ln(6)

    # --- 3. DETAILED EVIDENCE & AUDITOR NOTES ---
    pdf.chapter_title("3. Detailed Evidence & Auditor Notes")
    
    for i, claim in enumerate(claims, 1):
        c_status = getattr(claim, 'status', 'Unknown')
        is_verified = "Verified" in str(c_status)
        
        pdf.set_font('Helvetica', 'B', 11)
        if is_verified:
            pdf.set_text_color(0, 163, 224)
            status_label = "VERIFIED"
        else:
            pdf.set_text_color(220, 38, 38)
            status_label = "HALLUCINATION / UNVERIFIED"
            
        pdf.cell(0, 6, f"Claim {i}: [{status_label}]", 0, 1)
        
        # Claim Stated
        pdf.set_font('Helvetica', 'B', 10)
        pdf.set_text_color(0, 32, 91)
        pdf.cell(0, 5, "Claim Stated:", 0, 1)
        pdf.set_font('Helvetica', '', 9.5)
        pdf.set_text_color(50, 50, 50)
        pdf.multi_cell(0, 5, clean_pdf_text(getattr(claim, 'claim', '')))
        pdf.ln(1)
        
        # Core Feature & Limits
        pdf.set_font('Helvetica', 'B', 10)
        pdf.set_text_color(50, 50, 50)
        pdf.cell(40, 5, "Core Feature:", 0, 0)
        pdf.set_font('Helvetica', '', 9.5)
        pdf.cell(0, 5, clean_pdf_text(getattr(claim, 'core_policy_feature', 'N/A')), 0, 1)

        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(40, 5, "Stated Limit/Rule:", 0, 0)
        pdf.set_font('Helvetica', '', 9.5)
        pdf.cell(0, 5, clean_pdf_text(getattr(claim, 'stated_limit_or_rule', 'Standard terms')), 0, 1)

        # Client Application
        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(0, 5, "Client Application:", 0, 1)
        pdf.set_font('Helvetica', '', 9.5)
        pdf.multi_cell(0, 5, clean_pdf_text(getattr(claim, 'client_application', 'N/A')))
        pdf.ln(1)
        
        # Evidence Snippet
        pdf.set_font('Helvetica', 'B', 10)
        pdf.set_text_color(50, 50, 50)
        pdf.cell(0, 5, "Evidence Found in Policy Brochure:", 0, 1)
        pdf.set_font('Helvetica', 'I', 9)
        pdf.set_text_color(80, 80, 80)
        ev_snip = getattr(claim, 'evidence_snippet', 'No direct evidence found.')
        pdf.multi_cell(0, 4.5, f'"{clean_pdf_text(ev_snip)}"')
        pdf.ln(1)
        
        # Source Document
        pdf.set_font('Helvetica', 'B', 10)
        pdf.set_text_color(50, 50, 50)
        pdf.cell(40, 5, "Source Document:", 0, 0)
        pdf.set_font('Helvetica', '', 9.5)
        pdf.cell(0, 5, clean_pdf_text(getattr(claim, 'source_document', '')), 0, 1)
        pdf.ln(4)

    # --- 4. ADVISOR DECISION FRAMEWORK (Objective 2.2) ---
    pdf.add_page()
    pdf.chapter_title("4. Advisor Decision Framework")
    pdf.chapter_body(
        "Per Marsh compliance protocols, the Client Advisor must review this report before client distribution:\n\n"
        "1. APPROVE: If Overall Confidence >= 75% and zero claims are flagged as core feature Hallucinations.\n"
        "2. EDIT: If claims are flagged due to minor limit mismatches (e.g., INR 8,000 vs INR 10,000) or sub-limit phrasing. "
        "Manually correct the figure in the PPTX using the Evidence Snippet provided above.\n"
        "3. REJECT: If core policy features are hallucinated or the LLM applied benefits to a policy "
        "that does not offer them. Regenerate the pitch using stricter baseline documents."
    )

    pdf.output(output_path)
    logger.info(f"Generated Marsh compliance audit memo at: '{output_path}'")


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
        export_to_pptx(pitch_deck, output_path=str(pptx_path))

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

        # Render Corporate PDF Audit Report
        generate_audit_pdf(audit_report, company_name, selected_docs, str(pdf_path))

        logger.info(f"Generated all assets for '{company_name}' in '{OUTPUTS_DIR}'.")

        # Response payload
        return {
            "success": True,
            "company_name": company_name,
            "profile": profile.model_dump(),
            "pitch_deck": pitch_deck.model_dump(),
            "audit_report": audit_report.model_dump(),
            "retrieved_chunks": retrieved_chunks,
            "downloads": {
                "pptx": {
                    "filename": pptx_filename,
                    "url": f"/api/download/pptx/{pptx_filename}",
                },
                "pdf": {
                    "filename": pdf_filename,
                    "url": f"/api/download/pdf/{pdf_filename}",
                },
                "csv": {
                    "filename": csv_filename,
                    "url": f"/api/download/csv/{csv_filename}",
                },
                "json": {
                    "filename": json_filename,
                    "url": f"/api/download/json/{json_filename}",
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
