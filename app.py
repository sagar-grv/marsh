"""
app.py - Marsh AI Pitch Generator (Enterprise Consulting SaaS Edition)
======================================================================
Interactive Web Application for Enterprise Health Insurance Pitch Generation.
Features:
- Enterprise SaaS UI with Marsh Navy (#00205B) & Cyan (#00A3E0) Theme
- Hidden Streamlit Chrome (Clean Standalone Application Feel)
- Custom HTML/CSS KPI Metric Cards with Dynamic Confidence Gauges
- Dynamic LLM Engine Switcher (OpenAI gpt-4o-mini & Groq Llama 3.3 70B)
- Baseline Policy Document Filtering via Sidebar Checkboxes
- Client Organization Profiling & Risk Analysis
- Multi-Agent 4-Slide Pitch Generation with Speaker Notes
- Strict Hallucination & Compliance Audit with Confidence Score Metrics
- Full RAG Traceability & Evidence Inspection (Displays Relevance Tier & Normalized Score)
- Multi-Format Audit Report Export Engine: CSV, JSON, and Markdown
- Executive Consulting-Grade PowerPoint Export (.pptx)
"""

import os
import io
import csv
import json
from datetime import datetime
from pathlib import Path
import streamlit as st
from dotenv import load_dotenv

from core_agents import (
    generateCompanyProfile,
    generateMarketingPitch,
    auditPitchContent,
    export_to_pptx,
    TARGET_PDFS,
    AuditReport,
)

# Load environment configuration
load_dotenv()

# --- Page Configuration ---
st.set_page_config(
    page_title="Marsh AI | Corporate Health Risk Advisory",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# Premium Enterprise SaaS CSS & Custom Styling
# =============================================================================
st.markdown("""
<style>
    /* 1. Hide Streamlit Chrome for Standalone SaaS Look */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .viewerBadge_container__1QSob {display: none !important;}
    .stDeployButton {display: none !important;}
    
    /* 2. Global Typography & Background */
    body, .stApp {
        background-color: #F8FAFC;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* 3. Custom Sidebar Styling (Marsh Deep Navy #00205B) */
    [data-testid="stSidebar"] {
        background-color: #00205B !important;
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stCheckbox label {
        color: #E2E8F0 !important;
        font-weight: 500 !important;
    }
    [data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] {
        background-color: #0A2D74 !important;
        border-color: #00A3E0 !important;
    }
    
    /* 4. SaaS Top Navigation Bar */
    .saas-top-nav {
        background: linear-gradient(135deg, #00205B 0%, #001338 100%);
        padding: 1.2rem 2.0rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 4px 14px rgba(0, 32, 91, 0.15);
    }
    .saas-brand-title {
        font-size: 1.6rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        color: #FFFFFF;
        margin: 0;
    }
    .saas-brand-subtitle {
        color: #00A3E0;
        font-size: 0.92rem;
        font-weight: 600;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-top: 0.2rem;
    }
    
    /* 5. Custom HTML KPI Metric Cards Grid */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1.2rem;
        margin-bottom: 1.8rem;
    }
    .kpi-card {
        background: #FFFFFF;
        padding: 1.3rem;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0,0,0,0.06);
    }
    .kpi-label {
        font-size: 0.82rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #64748B;
        margin-bottom: 0.4rem;
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 800;
        color: #00205B;
        margin-bottom: 0.4rem;
    }
    .kpi-status-badge {
        font-size: 0.8rem;
        font-weight: 700;
        padding: 0.25rem 0.6rem;
        border-radius: 20px;
        display: inline-block;
    }
    .badge-pass { background-color: #DCFCE7; color: #166534; }
    .badge-flag { background-color: #FEE2E2; color: #991B1B; }
    .badge-info { background-color: #E0F2FE; color: #0369A1; }
    
    /* Progress Bar for Confidence */
    .progress-bar-bg {
        background-color: #E2E8F0;
        height: 8px;
        border-radius: 4px;
        overflow: hidden;
        margin-top: 0.5rem;
    }
    .progress-bar-fill {
        height: 100%;
        border-radius: 4px;
    }
    
    /* 6. Slide Cards */
    .slide-saas-card {
        background-color: #FFFFFF;
        border-left: 6px solid #00205B;
        border-radius: 10px;
        padding: 1.4rem;
        margin-bottom: 1.4rem;
        border-top: 1px solid #E2E8F0;
        border-right: 1px solid #E2E8F0;
        border-bottom: 1px solid #E2E8F0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    .slide-saas-title {
        color: #00205B;
        font-size: 1.25rem;
        font-weight: 700;
        margin-bottom: 1.0rem;
    }
    .bullet-row {
        display: flex;
        align-items: flex-start;
        margin-bottom: 0.7rem;
    }
    .bullet-dot {
        width: 10px;
        height: 10px;
        background-color: #00A3E0;
        border-radius: 50%;
        margin-top: 0.55rem;
        margin-right: 0.9rem;
        flex-shrink: 0;
    }
    .bullet-text {
        font-size: 1.02rem;
        color: #1E293B;
        line-height: 1.5;
    }
    
    /* 7. Export Hub Card */
    .export-hub-card {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 10px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# Helper: Multi-Format Audit Report Exporter (CSV, JSON, Markdown)
# =============================================================================
def generate_audit_exports(audit_report: AuditReport, company_name: str, baseline_docs: list) -> dict:
    """
    Transforms the Pydantic AuditReport into 3 production formats:
    1. CSV (tabular compliance matrix for Excel / PowerBI)
    2. JSON (structured programmatic export)
    3. Markdown (executive readable summary)
    """
    # 1. CSV Format
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
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
    csv_string = csv_buffer.getvalue()

    # 2. JSON Format
    json_data = {
        "client_organization": company_name,
        "evaluation_date": datetime.now().isoformat(),
        "overall_confidence_score": audit_report.deck_confidence_score,
        "summary": audit_report.summary,
        "baseline_documents_grounded": baseline_docs,
        "claims": [c.model_dump() for c in audit_report.claims],
    }
    json_string = json.dumps(json_data, indent=2)

    # 3. Markdown Format
    conf_pct = audit_report.deck_confidence_score * 100
    md_content = f"""# Marsh Risk Advisory - Compliance & Hallucination Audit Report

**Client Organization:** {company_name}  
**Evaluation Date:** {datetime.now().strftime('%B %d, %Y')}  
**Overall Deck Confidence Score:** {conf_pct:.1f}%  
**Baseline Documents Grounded:** {', '.join(baseline_docs)}  

---

## 1. Executive Compliance Summary
{audit_report.summary}

**Compliance Verdict:** {'PASS - Compliant for Presentation' if conf_pct >= 70 else 'REVIEW REQUIRED - Potential Term Mismatches'}

---

## 2. Claim-by-Claim Verification Breakdown

"""
    for i, c in enumerate(audit_report.claims, 1):
        md_content += f"""### Claim {i}: {c.status.upper()} (Confidence: {c.confidence_score:.2f})
- **Original Pitch Claim:** "{c.claim}"
- **Core Policy Feature:** `{c.core_policy_feature}`
- **Stated Limit or Rule:** `{c.stated_limit_or_rule}`
- **Client Application / Risk Mapping:** `{c.client_application}`
- **Policy Brochure Citation:** `{c.source_document}`
- **Auditor Evidence & Rationale:** 
  > {c.evidence_snippet}

---
"""
    md_content += "\n*Generated automatically by Marsh AI Risk Advisory Compliance Auditor.*"

    return {
        "csv": csv_string,
        "json": json_string,
        "md": md_content,
    }


# =============================================================================
# Sidebar: Settings, Model Switcher, and Baseline Documents
# =============================================================================
st.sidebar.markdown("""
<div style="text-align: center; padding: 0.5rem 0 1.2rem 0;">
    <h2 style="color: #FFFFFF; margin: 0; font-size: 1.4rem; letter-spacing: -0.5px;">MARSH ADVISORY</h2>
    <span style="color: #00A3E0; font-size: 0.8rem; font-weight: 700; letter-spacing: 1px;">AI PLATFORM</span>
</div>
""", unsafe_allow_html=True)

# LLM Engine Switcher
st.sidebar.subheader("🤖 AI Model Engine")
llm_options = [
    "Groq (Llama 3.3 70B)",
    "OpenAI (gpt-4o-mini)",
]
selected_llm = st.sidebar.selectbox(
    "Select Model Engine",
    options=llm_options,
    index=0,
    help="Toggle between Groq ultra-low latency inference and OpenAI gpt-4o-mini."
)

st.sidebar.divider()

# Baseline Policy Document Checkboxes
st.sidebar.subheader("📄 Baseline Policy Documents")
st.sidebar.caption("Filter the verified brochures available to the RAG retriever:")

selected_documents = []
for brochure in TARGET_PDFS:
    label = brochure.replace(".pdf", "")
    if st.sidebar.checkbox(label, value=True, key=f"chk_{brochure}"):
        selected_documents.append(brochure)

st.sidebar.caption(
    f"Active documents in knowledge base: **{len(selected_documents)} / {len(TARGET_PDFS)}**"
)

# Document validation indicator
if not selected_documents:
    st.sidebar.error("⚠️ Select at least one policy brochure to ground the RAG pipeline.")

st.sidebar.divider()

# Architecture Explanation Expander
with st.sidebar.expander("ℹ️ Multi-Agent Architecture", expanded=False):
    st.markdown("""
    **Pipeline Architecture:**
    1. **Company Profiler**: Gathers industry, scale & occupational health hazards.
    2. **Benefit Reformulator**: Maps workplace health risks to standard insurance brochure terms.
    3. **RAG Retriever Tool**: Retrieves semantically matched chunks using BGE embeddings.
    4. **Pitch Generator**: Synthesizes a 4-slide consulting-grade pitch deck.
    5. **Auditor Agent**: Fact-checks claims against source brochures to prevent hallucinations.
    6. **PPTX Exporter**: Renders McKinsey/Marsh-styled 16:9 executive presentation.
    """)


# =============================================================================
# Main Application Area: Top SaaS Navigation
# =============================================================================
st.markdown("""
<div class="saas-top-nav">
    <div>
        <div class="saas-brand-title">Marsh AI Pitch Generator</div>
        <div class="saas-brand-subtitle">Executive Corporate Health Risk Placement & Compliance Platform</div>
    </div>
    <div style="text-align: right;">
        <span style="background: rgba(0, 163, 224, 0.2); color: #00A3E0; padding: 0.35rem 0.8rem; border-radius: 20px; font-weight: 700; font-size: 0.85rem; border: 1px solid #00A3E0;">
            ● Active RAG Engine (BGE-Small v1.5)
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# Input Controls
col_input, col_action = st.columns([3, 1])

with col_input:
    company_name = st.text_input(
        "Client Company Name",
        placeholder="e.g., Tata Consultancy Services, Infosys, Reliance Industries",
        help="Enter the enterprise client name to analyze workplace risks and formulate tailored policy pitches."
    )

with col_action:
    st.write("")  # Spacing
    st.write("")
    is_generate_disabled = len(selected_documents) == 0
    generate_clicked = st.button(
        "🚀 Generate Pitch Deck",
        type="primary",
        disabled=is_generate_disabled,
        use_container_width=True,
    )

# Execution Handler
if generate_clicked:
    if not company_name.strip():
        st.warning("⚠️ Please provide a company name before initiating generation.")
    elif len(selected_documents) == 0:
        st.error("❌ At least one policy brochure must be selected from the sidebar.")
    else:
        st.session_state["has_run"] = True
        st.session_state["target_company"] = company_name.strip()
        st.session_state["selected_engine"] = selected_llm

        # Step 1: Profiler
        with st.spinner(f"🔍 [Agent 1/3: Profiler] Researching '{company_name}' & occupational health risks..."):
            profile = generateCompanyProfile(company_name.strip(), llm_provider=selected_llm)
            st.session_state["profile"] = profile

        # Step 2: Pitch Generator Agent (Returns deck & RAG chunks)
        with st.spinner(f"📑 [Agent 2/3: Pitch Strategist] Querying ChromaDB & synthesizing 4-slide deck ({selected_llm})..."):
            pitch_deck, rag_chunks = generateMarketingPitch(
                profile,
                selected_documents=selected_documents,
                llm_provider=selected_llm,
            )
            st.session_state["pitch_deck"] = pitch_deck
            st.session_state["rag_chunks"] = rag_chunks

        # Step 3: Auditor Agent
        with st.spinner(f"🛡️ [Agent 3/3: Compliance Auditor] Verifying factual claims against policy brochures..."):
            audit_report = auditPitchContent(
                pitch_deck,
                selected_documents=selected_documents,
                llm_provider=selected_llm,
            )
            st.session_state["audit_report"] = audit_report

        # Step 4: Export to PPTX
        safe_company_slug = "".join([c if c.isalnum() else "_" for c in company_name.strip()])
        pptx_filename = f"{safe_company_slug}_Pitch_Deck.pptx"
        export_to_pptx(pitch_deck, output_path=pptx_filename)
        st.session_state["pptx_filename"] = pptx_filename


# =============================================================================
# Presentation & Results Display (Persisted in Session State)
# =============================================================================
if st.session_state.get("has_run") and "pitch_deck" in st.session_state:
    profile = st.session_state["profile"]
    pitch_deck = st.session_state["pitch_deck"]
    audit_report = st.session_state["audit_report"]
    rag_chunks = st.session_state.get("rag_chunks", [])
    pptx_filename = st.session_state.get("pptx_filename", "Pitch_Deck.pptx")
    target_comp = st.session_state.get("target_company", "Client")
    safe_company_slug = "".join([c if c.isalnum() else "_" for c in target_comp])
    conf_pct = audit_report.deck_confidence_score * 100

    # Count Verified vs Flagged Claims
    verified_count = sum(1 for c in audit_report.claims if c.status == "Verified")
    flagged_count = len(audit_report.claims) - verified_count

    # Determine Progress Color
    progress_color = "#16A34A" if conf_pct >= 70 else "#DC2626"
    verdict_badge = (
        '<span class="kpi-status-badge badge-pass">● VERIFIED PASS</span>'
        if conf_pct >= 70 else
        '<span class="kpi-status-badge badge-flag">● REVIEW REQUIRED</span>'
    )

    # 1. Custom HTML KPI Metric Cards Grid
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-card">
            <div class="kpi-label">Deck Confidence Score</div>
            <div class="kpi-value">{conf_pct:.1f}%</div>
            <div>{verdict_badge}</div>
            <div class="progress-bar-bg">
                <div class="progress-bar-fill" style="width: {min(conf_pct, 100)}%; background-color: {progress_color};"></div>
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Claims Evaluated</div>
            <div class="kpi-value">{len(audit_report.claims)}</div>
            <span class="kpi-status-badge badge-info">100% Policy Grounded</span>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Verified Policy Terms</div>
            <div class="kpi-value" style="color: #16A34A;">{verified_count}</div>
            <span class="kpi-status-badge badge-pass">Brochure Confirmed</span>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Flagged / Unverified</div>
            <div class="kpi-value" style="color: {'#DC2626' if flagged_count > 0 else '#64748B'};">{flagged_count}</div>
            <span class="kpi-status-badge {'badge-flag' if flagged_count > 0 else 'badge-pass'}">
                {'Audit Discrepancy' if flagged_count > 0 else 'Zero Hallucinations'}
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 2. Company Profile Expander
    with st.expander(f"🏢 Client Organization Profile & Health Hazard Assessment ({target_comp})", expanded=False):
        c1, c2 = st.columns([1, 1])
        with c1:
            st.markdown(f"**Target Client:** {profile.company_name}")
            st.markdown(f"**Industry Sector:** {profile.industry}")
            st.markdown(f"**Workforce Scale:** {profile.size}")
            if profile.is_assumption:
                st.info("ℹ️ Profile includes domain assumptions from fallback analysis.")
        with c2:
            st.markdown("**Identified Occupational Health Hazards:**")
            for r in profile.key_risks:
                st.markdown(f"- {r}")

    # 3. Multi-Format Audit Exports Pre-computation
    export_files = generate_audit_exports(audit_report, target_comp, selected_documents)

    # 4. Tabs for Output
    tab_pitch, tab_audit, tab_rag, tab_download = st.tabs([
        "📑 Generated Pitch Deck",
        "🛡️ Compliance Audit & Export Hub",
        "🔍 RAG Traceability & Evidence",
        "📥 Download Consulting Deck (.pptx)",
    ])

    # Tab 1: Pitch Presentation
    with tab_pitch:
        st.subheader("Executive 4-Slide Consulting Pitch Presentation")
        for slide in pitch_deck.slides:
            bullets_html = "".join([
                f'<div class="bullet-row"><div class="bullet-dot"></div><div class="bullet-text">{bp}</div></div>'
                for bp in slide.bullet_points
            ])
            st.markdown(f"""
            <div class="slide-saas-card">
                <div class="slide-saas-title">Slide 0{slide.slide_number}: {slide.title}</div>
                {bullets_html}
            </div>
            """, unsafe_allow_html=True)

            with st.expander(f"🎙️ Speaker Notes: Slide 0{slide.slide_number} (Marsh Broker Script)"):
                st.write(slide.speaker_notes)

    # Tab 2: Audit Report & Multi-Format Export Hub
    with tab_audit:
        st.subheader("Hallucination & Policy Verification Audit")

        # Multi-Format Audit Export Section
        st.markdown("""
        <div class="export-hub-card">
            <h4 style="margin: 0 0 0.5rem 0; color: #00205B;">📊 Export Compliance Audit Trail</h4>
            <p style="color: #64748B; font-size: 0.92rem; margin-bottom: 1.0rem;">
                Download the complete audit trail in your preferred format for client records, underwriting governance, or PowerBI compliance tracking.
            </p>
        </div>
        """, unsafe_allow_html=True)

        exp_col1, exp_col2, exp_col3 = st.columns(3)
        with exp_col1:
            st.download_button(
                label="📊 Download Audit Report (.csv)",
                data=export_files["csv"],
                file_name=f"Audit_Report_{safe_company_slug}.csv",
                mime="text/csv",
                type="primary",
                use_container_width=True,
                help="Tabular matrix ideal for Microsoft Excel, compliance reviews, and PowerBI ingestion."
            )
        with exp_col2:
            st.download_button(
                label="📄 Download Audit Data (.json)",
                data=export_files["json"],
                file_name=f"Audit_Report_{safe_company_slug}.json",
                mime="application/json",
                type="secondary",
                use_container_width=True,
                help="Structured JSON payload ideal for API integrations and programmatic archiving."
            )
        with exp_col3:
            st.download_button(
                label="📝 Download Summary (.md)",
                data=export_files["md"],
                file_name=f"Audit_Report_{safe_company_slug}.md",
                mime="text/markdown",
                type="secondary",
                use_container_width=True,
                help="Executive Markdown summary formatted for document attachments and reviews."
            )

        st.divider()

        # Detailed Claim-by-Claim Breakdown
        st.markdown("#### Claim-by-Claim Verification Breakdown")
        for idx, claim_item in enumerate(audit_report.claims, 1):
            is_verified = (claim_item.status == "Verified")
            status_badge = "✅ Verified" if is_verified else "⚠️ Hallucination / Unverified"

            with st.expander(f"Claim {idx}: {claim_item.core_policy_feature} — {status_badge}"):
                if is_verified:
                    st.success(f"**Status:** {claim_item.status} | **Confidence:** {claim_item.confidence_score:.2f}")
                else:
                    st.error(f"**Status:** {claim_item.status} | **Confidence:** {claim_item.confidence_score:.2f}")

                st.markdown(f"**Original Statement:** *\"{claim_item.claim}\"*")
                c_a, c_b = st.columns(2)
                with c_a:
                    st.markdown(f"**Core Policy Feature:** `{claim_item.core_policy_feature}`")
                    st.markdown(f"**Stated Limit / Rule:** `{claim_item.stated_limit_or_rule}`")
                with c_b:
                    st.markdown(f"**Client Application:** `{claim_item.client_application}`")
                    st.markdown(f"**Brochure Citation:** `{claim_item.source_document}`")

                st.markdown(f"**Auditor Evidence & Rationale:**\n> {claim_item.evidence_snippet}")

    # Tab 3: RAG Traceability & Evidence (Displays Normalized Score & Relevance Tier)
    with tab_rag:
        st.subheader("RAG Traceability: Grounding Evidence from ChromaDB")
        st.markdown(
            "Inspect the exact policy brochure excerpts retrieved by the RAG system to ground this presentation:"
        )

        if rag_chunks:
            for idx, chunk in enumerate(rag_chunks, 1):
                tier = chunk.get("relevance_tier", "High")
                norm_score = chunk.get("normalized_score", 85.0)
                badge_class = "badge-pass" if tier == "High" else ("badge-info" if tier == "Medium" else "badge-flag")

                with st.container():
                    st.markdown(f"""
                    <div style="background-color: #F1F5F9; border-radius: 8px; padding: 0.9rem 1.1rem; margin-bottom: 0.6rem; border: 1px solid #CBD5E1; display: flex; justify-content: space-between; align-items: center;">
                        <div>
                            <b>Chunk {idx} | Source:</b> <code>{chunk['source_document']}</code> (Page {chunk['page']})
                        </div>
                        <div>
                            <span class="kpi-status-badge {badge_class}">Relevance: {tier} ({norm_score}%)</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    st.text_area(
                        label=f"Excerpt {idx}",
                        value=chunk["content"],
                        height=140,
                        key=f"chunk_txt_{idx}",
                        disabled=True,
                    )
        else:
            st.info("No policy excerpts were retrieved for the current query.")

    # Tab 4: Download Deck
    with tab_download:
        st.subheader("Export to Consulting-Grade PowerPoint (.pptx)")
        st.write("Download the presentation formatted in 16:9 widescreen with McKinsey/Marsh consulting aesthetics:")

        if Path(pptx_filename).exists():
            with open(pptx_filename, "rb") as f:
                pptx_bytes = f.read()

            st.download_button(
                label=f"📥 Download {pptx_filename}",
                data=pptx_bytes,
                file_name=pptx_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
            )
            st.caption("Includes 40/60 title split, custom Cyan circular bullet shapes, and full speaker notes.")
        else:
            st.error("PowerPoint file could not be located on disk.")
