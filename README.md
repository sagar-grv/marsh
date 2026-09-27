# Marsh AI Pitch Intelligence & Compliance Platform
## Enterprise Case Study Deliverable & System Guide

[![FastAPI](https://img.shields.io/badge/FastAPI-0.115+-009688.svg?style=flat&logo=fastapi)](https://fastapi.tiangolo.com)
[![ChromaDB](https://img.shields.io/badge/ChromaDB-bge--small--v1.5-blue.svg)](https://www.trychroma.com)
[![LangChain](https://img.shields.io/badge/LangChain-Multi--Agent-orange.svg)](https://python.langchain.com)
[![Compliance](https://img.shields.io/badge/Audit-Hallucination%20Firewall-emerald.svg)]()

Production-grade Multi-Agent RAG application built for **Marsh McLennan Corporate Health Advisory**. It automates client occupational risk research, policy clause matching, 4-slide consulting presentation generation, and independent regulatory compliance verification.

---

## 📋 Case Study Deliverables Summary

| Deliverable | Description | Primary Location |
| :--- | :--- | :--- |
| **1. Working Application** | FastAPI backend + Vanilla JS/Tailwind Corporate UI | [`server.py`](file:///d:/Qwen_marsh/server.py), [`frontend/`](file:///d:/Qwen_marsh/frontend) |
| **2. Sample Pitch Deck** | 16:9 McKinsey/Marsh executive PowerPoint deck | [`outputs/Tata_Consultancy_Services_Pitch_Deck.pptx`](file:///d:/Qwen_marsh/outputs/Tata_Consultancy_Services_Pitch_Deck.pptx) |
| **3. Audit Results (PDF)** | Formal corporate PDF compliance report with data matrix | [`outputs/Tata_Consultancy_Services_Audit_Report.pdf`](file:///d:/Qwen_marsh/outputs/Tata_Consultancy_Services_Audit_Report.pdf) |
| **4. Audit Results (CSV)** | Tabular compliance audit trail for Excel / PowerBI | [`outputs/Audit_Report_Tata_Consultancy_Services.csv`](file:///d:/Qwen_marsh/outputs/Audit_Report_Tata_Consultancy_Services.csv) |
| **5. Technical Write-up** | Full engineering architecture & design decision brief | [`WRITEUP.md`](file:///d:/Qwen_marsh/WRITEUP.md) |

---

## 🏛️ System Architecture Diagram

```
+--------------------------------------------------------------------------------------------------+
|                                1. INGESTION & KNOWLEDGE LAYER                                    |
|   Policy Brochures (HDFC, Niva Bupa, Care, ABHI) -> PyMuPDF Text -> 512-Token Recursive Chunks  |
|                     -> BAAI/bge-small-en-v1.5 -> Persistent ChromaDB Store                        |
+-------------------------------------------------+------------------------------------------------+
                                                  |
                                                  v
+------------------------+      +-----------------+-------------------+
| CLIENT COMPANY INPUT   | ---> | 2. COMPANY PROFILER AGENT           |
| e.g. "TCS", "Infosys"  |      |    DuckDuckGo Search + LLM Fallback |
+------------------------+      +-----------------+-------------------+
                                                  |
                                                  v
                                +-----------------+-------------------+
                                | 3. BENEFIT REFORMULATOR AGENT       |
                                |    Workplace Risks -> Policy Terms  |
                                +-----------------+-------------------+
                                                  |
                                                  v
                                +-----------------+-------------------+
                                | 4. NORMALIZED RAG RETRIEVER         |
                                |    BGE Prefix Query + Min-Max Score |
                                +-----------------+-------------------+
                                                  |
                                                  v
                                +-----------------+-------------------+
                                | 5. PITCH STRATEGIST AGENT           |
                                |    4-Slide Deck + Speaker Notes     |
                                +-----------------+-------------------+
                                                  |
                                                  v
                                +-----------------+-------------------+
                                | 6. COMPLIANCE AUDITOR AGENT         |
                                |    Feature Extraction & Verification|
                                +-----------------+-------------------+
                                                  |
                        +-------------------------+-------------------------+
                        |                                                   |
                        v                                                   v
+-----------------------+-----------------------+   +-----------------------+-----------------------+
| 7. EXPORT GENERATION SUITE                    |   | 8. ENTERPRISE ANTI-SLOP DASHBOARD             |
|   - 16:9 Widescreen Presentation (.pptx)      |   |   - Bloomberg/Stripe Flat Aesthetic           |
|   - Official Corporate Audit Report (.pdf)    |   |   - Real-Time KPI Cards (Score, Claims)       |
|   - Compliance Matrix (.csv)                  |   |   - HTML Claim Verification Table             |
|   - Programmatic JSON Audit Log (.json)       |   |   - Traceable Vector Passages View            |
+-----------------------------------------------+   +-----------------------------------------------+
```

---

## 🚀 Quickstart Guide

### 1. Environment Setup
Make sure Python 3.10+ is installed and your virtual environment is active:

```powershell
# Activate virtual environment
.\venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

Ensure `.env` in the project root contains your API keys:
```env
GROQ_API_KEY=your_groq_api_key_here
# Optional fallback
OPENAI_API_KEY=your_openai_api_key_here
```

### 2. Ingest Policy Documents
Rebuild the ChromaDB vector store using `bge-small-en-v1.5` with dense 512-character chunking:

```powershell
python ingest_policies.py
```

### 3. Launch the Application
Start the FastAPI server:

```powershell
python server.py
```

Open your browser and navigate to:
👉 **`http://127.0.0.1:8000`**

---

## 🔍 Key Multi-Agent Capabilities

### 1. Bridging the Semantic Gap (Relevance Optimization)
- **Problem**: Querying ChromaDB with client health risks (*"ergonomic strain", "eye fatigue", "stress"*) returned low scores (~0.31) because policy brochures use insurance terms (*"OPD", "AYUSH", "waiting period"*).
- **Solution**: The **Benefit Reformulator Agent** translates client health hazards into 6 to 10 standard policy provisions before querying the vector store, while preserving the raw risks for slide narrative generation.

### 2. Consulting vs. Hallucination Guard
- **Problem**: Naive auditors falsely flag valid consulting recommendations as hallucinations (e.g. recommending *AYUSH* for *workplace stress* because the word "stress" isn't in the brochure).
- **Solution**: The **Compliance Auditor Agent** decouples statutory terms (`core_policy_feature` and `stated_limit_or_rule`) from advisory applications (`client_application`), ensuring only fabricated limits or missing features are flagged.

### 3. Consulting-Grade Deliverables
- **16:9 Widescreen Presentations**: Generated via `python-pptx` using solid Marsh Navy `#00205B` covers and anchored text frames to eliminate overlapping text boxes.
- **Formal PDF Audit Reports**: Generated via `fpdf2` with solid Navy banners, executive summary statistics, and striped compliance tables.
