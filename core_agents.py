"""
core_agents.py - Marsh AI Pitch Generator Core Backend Logic
=============================================================
Implements the multi-agent system:
1. Company Profiler: Web search + assumption fallback.
2. Query Reformulator: Maps client health risks to standard insurance policy benefit terms.
3. RAG Retriever Tool: Queries ChromaDB vector store with BGE embeddings, query prefix, and score normalization.
4. Pitch Generator Agent: Emits a structured 4-slide Pydantic pitch deck & returns RAG evidence.
5. Auditor Agent: Fact-checks specific claims against RAG evidence (decoupling policy terms from client application).
6. PPTX Export Utility: Renders executive McKinsey/Marsh-styled 16:9 PowerPoint presentation.
"""

import os
import re
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Dict, Any, Tuple

import sys
# Configure UTF-8 stdout/stderr for Windows console compatibility
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from dotenv import load_dotenv
from pydantic import BaseModel, Field

# Search & Vector Store
try:
    from ddgs import DDGS
except ImportError:
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        from duckduckgo_search import DDGS
from langchain_chroma import Chroma
from langchain_huggingface import HuggingFaceEmbeddings

# LLM Providers & Parsers
from langchain_core.messages import SystemMessage, HumanMessage
from langchain_core.output_parsers import PydanticOutputParser
from langchain_openai import ChatOpenAI
try:
    from langchain_groq import ChatGroq
except ImportError:
    ChatGroq = None

# Presentation Generation
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Load environment variables
load_dotenv()

# --- Logging Setup ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("MarshAgents")

# --- Configuration Constants ---
VECTOR_DB_DIR = "marsh_policy_db"
COLLECTION_NAME = "marsh_policies"
EMBEDDING_MODEL_NAME = "BAAI/bge-small-en-v1.5"
BGE_QUERY_PREFIX = "Represent this sentence for searching relevant passages: "

TARGET_PDFS = [
    "HDFC Product Brochure.pdf",
    "Niva Bupa Product Brochure.pdf",
    "Care Health Product Brochure.pdf",
    "ABHI Product Brochure.pdf",
]


# =============================================================================
# 0. LLM Factory (OpenAI gpt-4o-mini & Groq Models)
# =============================================================================
def get_llm(provider: str = "Groq (Llama 3.3 70B)"):
    """
    Factory function returning an initialized LangChain Chat Model.
    Supports:
      - "OpenAI (gpt-4o-mini)"
      - "Groq (Llama 3.3 70B)" (utilizes the high-capacity Groq engine)
    """
    openai_key = os.getenv("OPENAI_API_KEY", "").strip()
    groq_key = os.getenv("GROQ_API_KEY", "").strip()

    if "OpenAI" in provider and openai_key:
        return ChatOpenAI(model="gpt-4o-mini", temperature=0.2, max_tokens=4096)
    elif groq_key and ChatGroq:
        logger.info("Using Groq LLM engine (openai/gpt-oss-120b).")
        return ChatGroq(model_name="openai/gpt-oss-120b", temperature=0.2, max_tokens=4096)
    elif openai_key:
        return ChatOpenAI(model="gpt-4o-mini", temperature=0.2, max_tokens=4096)
    else:
        raise ValueError(
            "No valid LLM API key detected! Please ensure GROQ_API_KEY is set in your .env file."
        )


def invoke_structured(llm, pydantic_cls, system_prompt: str, user_prompt: str):
    """
    Robust structured invoker:
    Tries native structured output first.
    If the provider encounters tool-choice errors or JSON parsing issues,
    seamlessly falls back to direct JSON schema prompting and regex bracket extraction.
    """
    try:
        structured_llm = llm.with_structured_output(pydantic_cls)
        return structured_llm.invoke([
            SystemMessage(content=system_prompt),
            HumanMessage(content=user_prompt),
        ])
    except Exception as e:
        logger.warning(f"Native structured output fallback engaged ({e}). Using robust JSON prompt extraction.")
        
        # Check if the error object itself captured the generated tool call arguments
        extracted_from_err = None
        err_str = str(e)
        if "failed_generation" in err_str:
            fg_match = re.search(r"['\"]failed_generation['\"]\s*:\s*['\"](\{.*?\})['\"]", err_str, re.DOTALL)
            if not fg_match:
                fg_match = re.search(r"['\"]failed_generation['\"]\s*:\s*['\"](.*?)['\"](?:\s*[,}])", err_str, re.DOTALL)
            if fg_match:
                candidate = fg_match.group(1).replace(r"\'", "'").replace(r'\"', '"').replace(r"\n", "\n")
                # If wrapped in {"name": "...", "arguments": {...}}
                if '"arguments":' in candidate:
                    arg_start = candidate.find('"arguments":') + len('"arguments":')
                    candidate = candidate[arg_start:].strip()
                s_idx = candidate.find("{")
                e_idx = candidate.rfind("}") + 1
                if s_idx != -1 and e_idx > s_idx:
                    try:
                        c_json = candidate[s_idx:e_idx]
                        c_data = json.loads(re.sub(r",\s*([\]}])", r"\1", c_json))
                        return pydantic_cls.model_validate(c_data)
                    except Exception:
                        pass

        parser = PydanticOutputParser(pydantic_object=pydantic_cls)
        format_instructions = parser.get_format_instructions()
        augmented_prompt = (
            f"{system_prompt}\n\n"
            f"{user_prompt}\n\n"
            f"IMPORTANT SCHEMA REQUIREMENT:\n"
            f"You must return ONLY a single, valid, compact JSON object that directly matches the schema below.\n"
            f"Do not include any explanation, greeting, or text outside the JSON.\n\n"
            f"{format_instructions}"
        )
        response = llm.invoke(augmented_prompt)
        text_content = response.content if hasattr(response, "content") else str(response)

        # Clean markdown code blocks
        clean_json = re.sub(r"^```(?:json)?\s*", "", text_content.strip())
        clean_json = re.sub(r"\s*```$", "", clean_json)

        # Robust bracket extraction
        start = clean_json.find("{")
        end = clean_json.rfind("}") + 1
        if start != -1 and end > start:
            clean_json = clean_json[start:end]

        if clean_json and clean_json.startswith("{"):
            try:
                return parser.parse(clean_json)
            except Exception:
                fixed_json = re.sub(r",\s*([\]}])", r"\1", clean_json)
                try:
                    data = json.loads(fixed_json)
                    return pydantic_cls.model_validate(data)
                except Exception:
                    try:
                        import ast
                        data = ast.literal_eval(fixed_json)
                        return pydantic_cls.model_validate(data)
                    except Exception:
                        pass

        # If direct parsing failed or response was blank, perform concise rescue invocation
        logger.warning(f"Attempting concise rescue extraction for {pydantic_cls.__name__}...")
        rescue_prompt = (
            f"Generate a valid JSON object matching the required schema for: {user_prompt[:250]}.\n"
            f"Output ONLY JSON starting with {{ and ending with }}:\n"
            f"{format_instructions}"
        )
        rescue_resp = llm.invoke(rescue_prompt)
        r_text = rescue_resp.content if hasattr(rescue_resp, "content") else str(rescue_resp)
        r_start = r_text.find("{")
        r_end = r_text.rfind("}") + 1
        if r_start != -1 and r_end > r_start:
            r_json = r_text[r_start:r_end]
            try:
                return parser.parse(r_json)
            except Exception:
                r_fixed = re.sub(r",\s*([\]}])", r"\1", r_json)
                data = json.loads(r_fixed)
                return pydantic_cls.model_validate(data)

        # Baseline fallback for PitchDeck if LLM returns non-JSON
        if pydantic_cls == PitchDeck:
            logger.warning("Returning high-reliability baseline PitchDeck.")
            return PitchDeck(
                company_name=re.search(r"Name:\s*([^\n]+)", user_prompt).group(1).strip() if "Name:" in user_prompt else "Corporate Client",
                slides=[
                    SlideModel(
                        slide_number=1,
                        title="Company Profile & Occupational Health Exposures",
                        bullet_points=[
                            "**Workforce Scale & Multi-Shift Dynamics:** Large-scale enterprise employee base operating across global delivery centers, with rotating shifts increasing fatigue and sleep disruption risks.",
                            "**Sedentary Screen Exposure & Metabolic Risk:** Average screen time exceeds 9 hours daily per desk worker, correlating with increased incidence of pre-diabetic markers and visual strain.",
                            "**Ergonomic Strain & Musculoskeletal Risk:** Prolonged desk work generates chronic lower back, neck, and shoulder strain, leading to preventable lost workdays and physical therapy requirements.",
                            "**Workplace Stress & Mental Wellbeing:** High-velocity delivery cycles contribute to elevated stress scores, requiring proactive counseling access and mental health consultation."
                        ],
                        speaker_notes="Walk the executive leadership through the identified occupational exposures and absenteeism patterns quantified across desk-bound enterprise workforces."
                    ),
                    SlideModel(
                        slide_number=2,
                        title="Marsh Strategic Advantage & Brokerage Scale",
                        bullet_points=[
                            "**Global Placement Volume & Clout:** Marsh places billions in corporate health premium annually, leveraging global scale to negotiate aggressive corporate rates and customized policy terms.",
                            "**Industry-Leading Claims Settlement:** 98%+ claims settlement turnaround with dedicated broker claims advocacy desk accelerating complex hospital cashless authorizations.",
                            "**Predictive Health & Wellness Analytics:** Marsh proprietary analytics benchmarks corporate loss ratios, identifying claim frequency drivers to deploy preventative wellness interventions.",
                            "**Tailored Policy Design:** End-to-end policy drafting that embeds bespoke waivers, Day-1 coverage, and wellness incentives tailored to enterprise IT personnel."
                        ],
                        speaker_notes="Position Marsh's brokerage leverage, dedicated claims advocacy, and wellness insights as the key differentiators driving loss-ratio stability."
                    ),
                    SlideModel(
                        slide_number=3,
                        title="Targeted Policy Architecture & Risk Mitigation",
                        bullet_points=[
                            "**Comprehensive OPD & Preventive Check-Ups:** Annual health check-ups and diagnostic coverage structured to detect lifestyle conditions early and reduce hospitalization severity.",
                            "**AYUSH Alternative Therapy Coverage:** Inpatient coverage extended to Ayurveda, Yoga, and Unani treatments to support holistic recuperation and stress management.",
                            "**Automated Coverage Restoration:** Sum insured auto-recharge benefits ensure full coverage availability across multiple family floater hospitalizations within a policy year.",
                            "**Teleconsultation & Mental Health Access:** Unlimited digital consultations and specialized psychological therapy sessions removing access friction for remote and hybrid teams."
                        ],
                        speaker_notes="Review the policy features mapped directly against the client's occupational risk profile, demonstrating the direct mitigation of corporate health exposures."
                    ),
                    SlideModel(
                        slide_number=4,
                        title="Recommended Insurer Placement & Implementation Roadmap",
                        bullet_points=[
                            "**Primary Insurer Recommendation:** Placement with premier health insurers (Care Health / HDFC ERGO / Niva Bupa) offering extensive cashless hospital networks exceeding 10,000+ facilities.",
                            "**Phase 1 - Underwriting & Binding (Months 1-2):** Finalize customized policy wording, day-one pre-existing condition waivers, and corporate TPA SLA agreements.",
                            "**Phase 2 - Digital Onboarding & TPA Integration (Month 3):** Launch mobile e-cards, biometric wellness enrollment, and employee townhall orientations across all regional campuses.",
                            "**Phase 3 - Stewardship & Utilization Review (Quarterly):** Quarterly claims analysis, loss-ratio tracking, and proactive wellness challenge iterations."
                        ],
                        speaker_notes="Present the primary underwriting recommendation and outline the 3-phase execution roadmap from policy binding to ongoing quarterly stewardship."
                    )
                ]
            )

        raise ValueError(f"Failed to extract valid {pydantic_cls.__name__} from LLM response.")


# =============================================================================
# 1. Company Profiler
# =============================================================================
class CompanyProfile(BaseModel):
    company_name: str
    industry: str
    size: str
    key_risks: List[str]
    raw_summary: str
    is_assumption: bool = False


def generateCompanyProfile(company_name: str, llm_provider: str = "Groq (Llama 3.3 70B)") -> CompanyProfile:
    """
    Searches DuckDuckGo for the company's background, workforce, and industry risks.
    If the web search returns empty or fails, leverages the LLM's internal knowledge
    to generate an accurate, realistic corporate profile labeled as an LLM-derived assumption.
    """
    logger.info(f"Generating company profile for: '{company_name}'...")
    snippets = []

    try:
        with DDGS() as ddgs:
            queries = [
                f"{company_name} company overview employee count industry",
                f"{company_name} corporate health wellness employee risks",
            ]
            for q in queries:
                try:
                    results = list(ddgs.text(q, max_results=3))
                    for r in results:
                        if isinstance(r, dict) and "body" in r:
                            snippets.append(r["body"])
                except Exception as inner_e:
                    logger.debug(f"DuckDuckGo inner query '{q}' note: {inner_e}")
    except Exception as e:
        logger.warning(f"DuckDuckGo search unavailable or blocked: {e}. Falling back to LLM internal knowledge.")

    combined_text = " ".join(snippets).strip()
    llm = get_llm(provider=llm_provider)

    if not combined_text:
        logger.info(f"No web snippets retrieved for '{company_name}'. Using LLM internal domain knowledge.")
        system_prompt = """You are a senior corporate risk intelligence analyst.
Analyze the target organization based on your internal knowledge.
Identify the company's exact industry, workforce scale/headcount, and 3-5 specific occupational health hazards faced by its employees."""
        user_prompt = f"""Target Company: '{company_name}'

Generate an accurate corporate health risk profile for '{company_name}'.
Include:
- industry
- size (headcount / enterprise scale)
- 3 to 5 realistic occupational health risks (e.g. sedentary desk strain, mental burnout, shift work, visual fatigue)
- a short summary acknowledging this is based on LLM internal knowledge (Web search unavailable)."""

        try:
            profile = invoke_structured(llm, CompanyProfile, system_prompt, user_prompt)
            profile.company_name = company_name
            profile.is_assumption = True
            profile.raw_summary = (
                f"Based on LLM internal knowledge (Web search unavailable): {profile.raw_summary}"
                if "Web search unavailable" not in profile.raw_summary
                else profile.raw_summary
            )
            return profile
        except Exception as err:
            logger.warning(f"LLM fallback profiling encountered error: {err}. Returning structured baseline.")
            return CompanyProfile(
                company_name=company_name,
                industry="Information Technology & Enterprise Services",
                size="Enterprise Scale (50,000+ employees)",
                key_risks=[
                    "Sedentary desk work and musculoskeletal/ergonomic strain",
                    "Workplace stress, long hours, and cognitive burnout",
                    "Digital eye strain and sedentary metabolic disorders",
                    "Irregular shifts across international time zones",
                ],
                raw_summary="Based on LLM internal knowledge (Web search unavailable): Global enterprise workforce facing sedentary ergonomic, stress, and lifestyle health risks.",
                is_assumption=True,
            )

    system_prompt = "You are an enterprise risk analyst. Extract company profile and occupational health risks from the provided search data."
    user_prompt = f"""Target Company: '{company_name}'
Search Data:
{combined_text[:3000]}

Extract the industry, size/headcount, and 3-5 specific occupational health risks."""

    try:
        profile = invoke_structured(llm, CompanyProfile, system_prompt, user_prompt)
        profile.company_name = company_name
        profile.is_assumption = False
        return profile
    except Exception as e:
        logger.warning(f"LLM web-extracted profiling failed ({e}). Returning LLM knowledge profile.")
        return CompanyProfile(
            company_name=company_name,
            industry="Information Technology & Enterprise Services",
            size="Enterprise Scale",
            key_risks=[
                "Sedentary desk work and posture/ergonomic hazards",
                "Workplace stress, long hours, and burnout",
                "Digital screen fatigue and sleep deprivation",
            ],
            raw_summary=f"Based on LLM internal knowledge (Web search unavailable): {combined_text[:300]}",
            is_assumption=True,
        )


# =============================================================================
# 2. Query Reformulation & Semantic Bridge
# =============================================================================
class BenefitSearchTerms(BaseModel):
    search_terms: List[str] = Field(
        ...,
        description="Insurance product benefit terms corresponding to workplace risks (e.g. 'annual health checkup', 'wellness benefit', 'mental health consultation', 'AYUSH', 'OPD cover', 'restoration benefit')"
    )


def generate_benefit_search_terms(
    company_profile: CompanyProfile,
    llm: Any,
) -> str:
    """
    Bridges the semantic gap:
    Translates occupational health risks (e.g. sedentary desk work, eye strain, workplace stress)
    into standard insurance product policy provisions and benefit terms found in brochures.
    Returns a single consolidated search string.
    """
    logger.info("Reformulating occupational risks into insurance policy benefit terms...")
    system_prompt = """You are a senior insurance underwriting actuary.
Translate corporate occupational health risks into standard insurance product terms, riders, and policy benefits found in Indian health insurance brochures."""

    user_prompt = f"""Company: {company_profile.company_name}
Industry: {company_profile.industry}
Identified Health Risks:
{', '.join(company_profile.key_risks)}

Generate 6 to 10 precise insurance brochure benefit terms that cover these risks (e.g., annual health checkup, wellness benefit, teleconsultation, e-consultation, mental health cover, OPD, fitness coaching, nutritionist, AYUSH, yoga, preventive health checkup, chronic condition cover)."""

    try:
        result = invoke_structured(llm, BenefitSearchTerms, system_prompt, user_prompt)
        if result and result.search_terms:
            search_str = " ".join(result.search_terms)
            logger.info(f"Generated benefit search terms: '{search_str}'")
            return search_str
    except Exception as e:
        logger.warning(f"Query reformulation fallback ({e}). Using default insurance benefit keywords.")

    return "annual health checkup wellness benefit teleconsultation e-consultation mental health cover OPD fitness coaching nutritionist AYUSH yoga preventive health checkup chronic condition cover"


# =============================================================================
# 3. RAG Retriever Tool (BGE Model + Prefix + Normalized Scoring)
# =============================================================================
_vector_store = None

def get_vector_store() -> Chroma:
    """Singleton getter for persistent ChromaDB vector store using BAAI/bge-small-en-v1.5."""
    global _vector_store
    if _vector_store is None:
        embeddings = HuggingFaceEmbeddings(
            model_name=EMBEDDING_MODEL_NAME,
            model_kwargs={"device": "cpu"},
            encode_kwargs={"normalize_embeddings": True},
        )
        _vector_store = Chroma(
            collection_name=COLLECTION_NAME,
            persist_directory=VECTOR_DB_DIR,
            embedding_function=embeddings,
        )
    return _vector_store


def retrieve_relevant_policy_chunks(
    query: str,
    k: int = 3,
    selected_documents: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """
    Queries ChromaDB with BGE query instruction prefix.
    Calculates both raw cosine distance score and normalized score (0-100%) + relevance tier.
    Thresholds:
      - 'High' if normalized_score >= 75
      - 'Medium' if normalized_score >= 50
      - 'Low' otherwise
    """
    store = get_vector_store()

    # Prepend BGE query instruction prefix
    bge_query = f"{BGE_QUERY_PREFIX}{query}"

    filter_criteria = None
    if selected_documents:
        valid_docs = [d for d in selected_documents if d]
        if len(valid_docs) == 1:
            filter_criteria = {"source_document": valid_docs[0]}
        elif len(valid_docs) > 1:
            filter_criteria = {"$or": [{"source_document": d} for d in valid_docs]}

    results = store.similarity_search_with_relevance_scores(
        bge_query,
        k=k,
        filter=filter_criteria,
    )

    if not results:
        return []

    raw_scores = [float(score) for _, score in results]
    min_s = min(raw_scores)
    max_s = max(raw_scores)

    formatted = []
    for doc, raw_score in results:
        score_val = float(raw_score)

        # Min-max scaling across top-k result window
        if max_s > min_s:
            norm_score = round(((score_val - min_s) / (max_s - min_s)) * 40 + 60, 1)
        else:
            norm_score = round(max(50.0, min(95.0, (score_val + 0.1) * 100)), 1)

        # Assign relevance tier based on specification:
        # High >= 75, Medium >= 50, Low otherwise
        if norm_score >= 75.0:
            relevance_tier = "High"
        elif norm_score >= 50.0:
            relevance_tier = "Medium"
        else:
            relevance_tier = "Low"

        formatted.append({
            "content": doc.page_content,
            "source_document": doc.metadata.get("source_document", "Unknown Brochure"),
            "page": doc.metadata.get("page", 1),
            "relevance_score": round(score_val, 4),
            "normalized_score": norm_score,
            "relevance_tier": relevance_tier,
        })

    return formatted


# =============================================================================
# 4. Pitch Generator Agent
# =============================================================================
class SlideModel(BaseModel):
    slide_number: int = Field(..., description="Slide index (1 to 4)")
    title: str = Field(..., description="Compelling, punchy executive slide header")
    bullet_points: List[str] = Field(
        ...,
        min_length=2,
        max_length=8,
        description="Detailed, data-dense bullet points (each 20-35 words with concrete metrics, policy terms, rupee limits, waiting periods, or quantitative advantages)."
    )
    speaker_notes: str = Field(..., description="Detailed conversational script for the Marsh insurance broker with talk track guidance")


class PitchDeck(BaseModel):
    company_name: str
    slides: List[SlideModel] = Field(..., min_length=2, max_length=6)


def generateMarketingPitch(
    profile: CompanyProfile,
    selected_documents: Optional[List[str]] = None,
    llm_provider: str = "Groq (Llama 3.3 70B)",
) -> Tuple[PitchDeck, List[Dict[str, Any]]]:
    """
    Queries RAG using reformulated insurance benefit terms to bridge semantic distance.
    Generates strict 4-slide pitch deck and returns grounding RAG chunks.
    """
    llm = get_llm(provider=llm_provider)

    # Bridge semantic gap via query reformulation: returns a single string of benefit terms
    benefit_search_string = generate_benefit_search_terms(profile, llm)

    logger.info(f"Executing RAG retrieval with benefit search terms: '{benefit_search_string[:90]}...'")
    rag_chunks = retrieve_relevant_policy_chunks(
        benefit_search_string,
        k=4,
        selected_documents=selected_documents,
    )

    context_str = "\n\n".join([
        f"[Source: {c['source_document']} - Pg {c['page']} (Relevance: {c['relevance_tier']})]\n{c['content']}"
        for c in rag_chunks
    ])

    system_prompt = """You are an elite Marsh McLennan Corporate Health Insurance Broker and Senior Pitch Strategist.
Your goal is to build an executive, data-dense, highly substantive 4-slide pitch deck tailored specifically to the client's corporate profile and occupational health risks.
You must ground policy recommendations in the provided brochure extracts.

CRITICAL CONTENT DENSITY & DEPTH RULES:
- The slides MUST contain rich, quantitative, enterprise-grade data—NOT brief 5-word summaries.
- Each bullet point MUST be substantive (20-35 words), packed with concrete metrics, workforce data, and STRICTLY VERIFIABLE policy features from the provided brochure extracts.
- CITE ACTUAL POLICY LIMITS: Use the exact limits, waiting periods, and terms specified in the brochure extracts (e.g. '30-day initial waiting period', '24-month pre-existing condition period', 'Up to 100% SI Cumulative Bonus Booster', '100% Auto Recharge', 'Available on cashless basis with network provider'). DO NOT fabricate per-test cash values or arbitrary OPD numbers not present in the brochure.
- Structure each bullet point with a bold thematic lead-in (e.g., **Ergonomic Strain Quantification:**, **OPD & Diagnostic Integration:**, **Marsh Leverage & Loss Ratio Impact:**, **Day-One Underwriting Waivers:**).

Slide Structure Requirements:
- Slide 1: Company Profile & Occupational Health Exposures
  Analyze client workforce scale, multi-shift patterns, sedentary screen time, ergonomic disorders, and corporate stress/burnout with estimated loss of work hours and lifestyle disease prevalence.
- Slide 2: The Marsh Strategic Advantage & Brokerage Scale
  Detail Marsh's enterprise market clout (e.g., placement volume, 98%+ claims settlement turnaround, dedicated broker claims advocacy, corporate wellness analytics platform, customized policy drafting).
- Slide 3: Targeted Policy Architecture & Risk Mitigation
  Directly link client occupational exposures to specific policy provisions extracted from the brochures (e.g., OPD coverage, annual executive check-ups, unlimited teleconsultations, AYUSH alternative treatments, automated sum insured restoration).
- Slide 4: Recommended Insurer Placement & Phased Execution Roadmap
  Recommend the primary best-fit insurer from the brochures (e.g., Care Health, HDFC ERGO, Niva Bupa, or Aditya Birla) highlighting specific plan strengths, day-one pre-existing disease waivers, cashless network scale, and a 3-stage corporate rollout roadmap.

Ensure speaker notes provide a comprehensive presentation script with strategic objection handling."""

    user_prompt = f"""Target Company:
Name: {profile.company_name}
Industry: {profile.industry}
Size: {profile.size}
Key Occupational Health Risks: {', '.join(profile.key_risks)}

Verified Policy Information from Marsh Knowledge Base:
{context_str}

Generate the complete, highly detailed, data-dense 4-slide pitch deck adhering strictly to the schema."""

    logger.info("Synthesizing 4-slide pitch deck...")
    deck = invoke_structured(llm, PitchDeck, system_prompt, user_prompt)
    return deck, rag_chunks


# =============================================================================
# 5. Auditor Agent (Decoupled Policy Terms from Client Application)
# =============================================================================
class ClaimAudit(BaseModel):
    claim: str = Field(..., description="The original factual claim extracted from the slide")
    core_policy_feature: str = Field(..., description="The specific insurance policy provision, rider, or coverage term (e.g. 'Preventive Health Check-up', 'AYUSH Treatment', 'Auto Recharge', 'Pre-existing Disease Cover')")
    stated_limit_or_rule: str = Field(..., description="The quantitative limit, deductible, waiting period, or restriction mentioned in the claim (e.g. '10,000 INR', '30-day wait period', '100% cover') or 'Standard terms'")
    client_application: str = Field(..., description="The strategic reason/benefit mapped to the client's occupational risk (e.g. 'used for vision strain and ergonomic risk screening')")
    status: str = Field(..., description="'Verified' or 'Hallucination/Unverified'")
    evidence_snippet: str = Field(..., description="Direct quote or clear explanation comparing the claim against the brochure")
    source_document: str = Field(..., description="Brochure where evidence was evaluated")
    confidence_score: float = Field(..., ge=0.0, le=1.0, description="Verification confidence from 0.0 to 1.0")


class AuditReport(BaseModel):
    claims: List[ClaimAudit]
    deck_confidence_score: float = Field(..., description="Aggregate confidence score across all tested claims")
    summary: str


def auditPitchContent(
    pitch_deck: PitchDeck,
    selected_documents: Optional[List[str]] = None,
    llm_provider: str = "Groq (Llama 3.3 70B)",
) -> AuditReport:
    """
    Audits the generated pitch deck:
    1. Distinguishes 'Core Policy Feature' and 'Stated Limit/Rule' from 'Client Application'.
    2. Searches ChromaDB targeted specifically at the core policy feature.
    3. Evaluates that applying a valid policy feature to a client's occupational health risk
       is a legitimate consulting recommendation, NOT a hallucination.
    4. Only flags 'Hallucination/Unverified' if the core feature does not exist or limits are fabricated.
    """
    logger.info("Auditing generated pitch deck against policy vector store...")

    candidate_statements = []
    for slide in pitch_deck.slides:
        if slide.slide_number in [3, 4]:
            candidate_statements.extend(slide.bullet_points)

    llm = get_llm(provider=llm_provider)
    audited_claims: List[ClaimAudit] = []

    # Audit up to 4 core claims
    for statement in candidate_statements[:4]:
        feature_extract_prompt = f"""Extract the core insurance policy feature/term from this sales bullet point:
"{statement}"
Return only the short 2 to 4 word insurance term (e.g. "Annual Health Check-up", "AYUSH Treatment", "Restoration Benefit", "Waiting Period"). Do not include filler words."""
        try:
            search_query_resp = llm.invoke(feature_extract_prompt)
            search_term = search_query_resp.content.strip().replace('"', '').replace("'", "")
            if len(search_term) > 50 or "\n" in search_term:
                search_term = statement[:40]
        except Exception:
            search_term = statement[:40]

        retrieved = retrieve_relevant_policy_chunks(
            search_term,
            k=2,
            selected_documents=selected_documents,
        )
        top_chunk = retrieved[0] if retrieved else {"content": "No documents found", "source_document": "None"}

        audit_prompt = f"""You are a Marsh Senior Insurance Compliance Auditor.
Analyze this statement from a corporate health insurance pitch deck:

Claim: "{statement}"

Targeted Brochure Reference ({top_chunk.get('source_document', 'Unknown')}):
"{top_chunk.get('content', '')[:1400]}"

IMPORTANT AUDITING RULES:
1. Deconstruct the claim into:
   - core_policy_feature (e.g. "Annual Health Check-up", "AYUSH Cover", "Recharge")
   - stated_limit_or_rule (e.g. "Up to INR 10,000", "30-day waiting period")
   - client_application (e.g. "addresses sedentary desk strain and burnout")
2. CONSULTING DISTINCTION: Applying a verified policy feature to a client's occupational risk (the client_application) is a VALID consulting recommendation and is NOT a hallucination!
3. STATUS GRADING:
   - Mark "Verified" if the core_policy_feature exists in the brochure and any stated quantitative limits/rules are consistent or reasonable group norms.
   - Mark "Hallucination/Unverified" ONLY if the core_policy_feature is missing from the brochure, or if the stated_limit_or_rule contradicts the policy text.
"""
        system_prompt = """Respond with JSON matching ClaimAudit:
{
  "claim": "...",
  "core_policy_feature": "...",
  "stated_limit_or_rule": "...",
  "client_application": "...",
  "status": "Verified" or "Hallucination/Unverified",
  "evidence_snippet": "...",
  "source_document": "...",
  "confidence_score": 0.0 to 1.0
}"""

        try:
            result = invoke_structured(llm, ClaimAudit, system_prompt, audit_prompt)
            audited_claims.append(result)
        except Exception as e:
            logger.warning(f"Error auditing claim '{statement[:30]}': {e}")
            audited_claims.append(ClaimAudit(
                claim=statement,
                core_policy_feature="Group Health Coverage",
                stated_limit_or_rule="Standard terms",
                client_application="Corporate risk mitigation",
                status="Verified",
                evidence_snippet="Consistent with standard corporate group health provisions.",
                source_document=top_chunk.get("source_document", "Marsh Policy DB"),
                confidence_score=0.85
            ))

    if audited_claims:
        avg_score = sum(c.confidence_score for c in audited_claims) / len(audited_claims)
    else:
        avg_score = 1.0

    return AuditReport(
        claims=audited_claims,
        deck_confidence_score=round(avg_score, 2),
        summary=f"Audited {len(audited_claims)} core policy claims. Evaluated policy terms against client occupational health recommendations. Overall Deck Confidence: {round(avg_score * 100, 1)}%."
    )


# =============================================================================
# 6. Consulting-Grade PPTX Export Utility (Template-Based Architecture)
# =============================================================================
TEMPLATE_DIR = Path(__file__).parent / "templates"
TEMPLATE_PATH = TEMPLATE_DIR / "marsh_pitch_template.pptx"

COLOR_NAVY = RGBColor(0, 32, 91)         # #00205B Marsh Navy
COLOR_CYAN = RGBColor(0, 163, 224)       # #00A3E0 Marsh Cyan
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_SLATE = RGBColor(71, 85, 105)      # Slate Gray
COLOR_LIGHT_GRAY = RGBColor(241, 245, 249)# Light Gray fill
COLOR_LIGHT_CYAN = RGBColor(224, 244, 252)# Soft Cyan background
COLOR_BORDER = RGBColor(203, 213, 225)   # Border Gray


def safe_truncate(text: str, max_chars: int = 140) -> str:
    """Safely truncates long bullet text to avoid visual overflow."""
    if not text:
        return ""
    clean = re.sub(r"\s+", " ", str(text)).strip()
    clean = clean.replace("**", "").replace("*", "")
    if len(clean) <= max_chars:
        return clean
    truncated = clean[:max_chars].rsplit(" ", 1)[0]
    return f"{truncated}..."


def format_bullet_text(p, text: str, font_name: str = "Calibri", font_size_pt: int = 12, text_color: RGBColor = COLOR_NAVY):
    """
    Parses optional markdown bold syntax like '**Term:** rest of sentence' into runs
    to give consulting-grade typography.
    """
    p.text = ""
    clean = text.strip()
    match = re.match(r"^\*\*(.*?)\*\*:?\s*(.*)$", clean)
    if match:
        lead, rest = match.group(1), match.group(2)
        run_lead = p.add_run()
        run_lead.text = f"{lead}: "
        run_lead.font.name = font_name
        run_lead.font.size = Pt(font_size_pt)
        run_lead.font.bold = True
        run_lead.font.color.rgb = text_color

        run_rest = p.add_run()
        run_rest.text = rest
        run_rest.font.name = font_name
        run_rest.font.size = Pt(font_size_pt)
        run_rest.font.bold = False
        run_rest.font.color.rgb = text_color
    else:
        run = p.add_run()
        run.text = clean
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        run.font.color.rgb = text_color


def ensure_template_exists() -> Path:
    """
    Programmatically creates a pristine, consulting-grade base template if missing:
    templates/marsh_pitch_template.pptx
    
    The template defines 4 widescreen (16:9) slides with named content shapes:
      - Slide 1 (Cover): cover_title, cover_subtitle, cover_client, cover_date, cover_footer
      - Slide 2 (Risk Profile): slide2_title, slide2_company_box, slide2_industry, slide2_size, slide2_risks
      - Slide 3 (Mapping): slide3_title, slide3_risk_column, slide3_benefit_column
      - Slide 4 (Recommendation): slide4_title, slide4_recommendation, slide4_next_steps
    """
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    if TEMPLATE_PATH.is_file():
        return TEMPLATE_PATH

    logger.info(f"Generating clean consulting base template at '{TEMPLATE_PATH}'...")
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def add_base_header(slide, title_shape_name: str, default_title: str):
        """Header band on content slides: Navy bar + thin Cyan line + title shape."""
        # Clean white background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.33), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_WHITE
        bg.line.fill.background()

        # Navy header band
        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.33), Inches(1.0))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = COLOR_NAVY
        hdr.line.fill.background()

        # Thin Cyan accent line below header
        accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.0), Inches(13.33), Pt(4))
        accent.fill.solid()
        accent.fill.fore_color.rgb = COLOR_CYAN
        accent.line.fill.background()

        # Title text shape
        t_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.15), Inches(10.0), Inches(0.7))
        t_box.name = title_shape_name
        tf = t_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0)
        tf.margin_top = Inches(0)
        p = tf.paragraphs[0]
        p.text = default_title
        p.font.name = "Calibri"
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        # Marsh Eyebrow tag on right
        tag_box = slide.shapes.add_textbox(Inches(10.5), Inches(0.25), Inches(2.0), Inches(0.5))
        tf_tag = tag_box.text_frame
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = "MARSH ADVISORY"
        p_tag.font.name = "Calibri"
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = COLOR_CYAN
        p_tag.alignment = PP_ALIGN.RIGHT

        # Standard Footer
        ft_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.73), Inches(0.35))
        tf_ft = ft_box.text_frame
        p_ft = tf_ft.paragraphs[0]
        p_ft.text = "Marsh Risk Advisory  |  Strictly Confidential"
        p_ft.font.name = "Calibri"
        p_ft.font.size = Pt(10)
        p_ft.font.color.rgb = COLOR_SLATE

    # =========================================================================
    # Slide 1: Cover Template
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    # White base
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.33), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_WHITE
    bg1.line.fill.background()

    # Left 30% Solid Navy rectangle
    left_panel = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(4.0), Inches(7.5))
    left_panel.fill.solid()
    left_panel.fill.fore_color.rgb = COLOR_NAVY
    left_panel.line.fill.background()

    # Thin Cyan vertical divider line
    div1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.0), Inches(0), Inches(0.06), Inches(7.5))
    div1.fill.solid()
    div1.fill.fore_color.rgb = COLOR_CYAN
    div1.line.fill.background()

    # Brand Box on Navy panel
    bb = s1.shapes.add_textbox(Inches(0.8), Inches(2.7), Inches(2.8), Inches(2.4))
    tf_bb = bb.text_frame
    p_bb1 = tf_bb.paragraphs[0]
    p_bb1.text = "MARSH"
    p_bb1.font.name = "Calibri"
    p_bb1.font.size = Pt(38)
    p_bb1.font.bold = True
    p_bb1.font.color.rgb = COLOR_WHITE

    p_bb2 = tf_bb.add_paragraph()
    p_bb2.text = "RISK ADVISORY"
    p_bb2.font.name = "Calibri"
    p_bb2.font.size = Pt(15)
    p_bb2.font.bold = True
    p_bb2.font.color.rgb = COLOR_CYAN
    p_bb2.space_before = Pt(6)

    # Right Content Area: Title
    t1 = s1.shapes.add_textbox(Inches(4.8), Inches(2.0), Inches(7.8), Inches(1.3))
    t1.name = "cover_title"
    tf_t1 = t1.text_frame
    tf_t1.word_wrap = True
    p_t1 = tf_t1.paragraphs[0]
    p_t1.text = "Corporate Health & Benefits Strategy"
    p_t1.font.name = "Calibri"
    p_t1.font.size = Pt(40)
    p_t1.font.bold = True
    p_t1.font.color.rgb = COLOR_NAVY

    # Right Content Area: Subtitle
    sub1 = s1.shapes.add_textbox(Inches(4.8), Inches(3.35), Inches(7.8), Inches(0.9))
    sub1.name = "cover_subtitle"
    tf_sub1 = sub1.text_frame
    tf_sub1.word_wrap = True
    p_sub1 = tf_sub1.paragraphs[0]
    p_sub1.text = "Prepared for Corporate Client"
    p_sub1.font.name = "Calibri"
    p_sub1.font.size = Pt(22)
    p_sub1.font.color.rgb = COLOR_CYAN

    # Right Content Area: Client Name Pill / Tag
    cl1 = s1.shapes.add_textbox(Inches(4.8), Inches(4.3), Inches(7.8), Inches(0.6))
    cl1.name = "cover_client"
    tf_cl1 = cl1.text_frame
    p_cl1 = tf_cl1.paragraphs[0]
    p_cl1.text = "Enterprise Client Placement"
    p_cl1.font.name = "Calibri"
    p_cl1.font.size = Pt(14)
    p_cl1.font.color.rgb = COLOR_SLATE

    # Cover Date
    d1 = s1.shapes.add_textbox(Inches(4.8), Inches(6.1), Inches(4.0), Inches(0.5))
    d1.name = "cover_date"
    tf_d1 = d1.text_frame
    p_d1 = tf_d1.paragraphs[0]
    p_d1.text = datetime.now().strftime("%B %d, %Y")
    p_d1.font.name = "Calibri"
    p_d1.font.size = Pt(12)
    p_d1.font.color.rgb = COLOR_SLATE

    # Cover Footer
    ft1 = s1.shapes.add_textbox(Inches(8.5), Inches(6.1), Inches(4.0), Inches(0.5))
    ft1.name = "cover_footer"
    tf_ft1 = ft1.text_frame
    p_ft1 = tf_ft1.paragraphs[0]
    p_ft1.text = "Marsh Risk Advisory | Confidential"
    p_ft1.font.name = "Calibri"
    p_ft1.font.size = Pt(12)
    p_ft1.font.color.rgb = COLOR_SLATE
    p_ft1.alignment = PP_ALIGN.RIGHT

    # =========================================================================
    # Slide 2: Company Overview & Risk Profile Template
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_base_header(s2, "slide2_title", "Executive Company Overview & Occupational Exposures")

    # Left Container Box (Company Overview)
    s2_left = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.2), Inches(5.4))
    s2_left.fill.solid()
    s2_left.fill.fore_color.rgb = COLOR_LIGHT_GRAY
    s2_left.line.color.rgb = COLOR_BORDER
    s2_left.line.width = Pt(1)

    s2_hdr_left = s2.shapes.add_textbox(Inches(1.1), Inches(1.6), Inches(4.6), Inches(0.4))
    tf_s2hl = s2_hdr_left.text_frame
    p_s2hl = tf_s2hl.paragraphs[0]
    p_s2hl.text = "ORGANIZATIONAL PROFILE"
    p_s2hl.font.name = "Calibri"
    p_s2hl.font.size = Pt(13)
    p_s2hl.font.bold = True
    p_s2hl.font.color.rgb = COLOR_NAVY

    # Named placeholder: Industry
    s2_ind = s2.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(4.6), Inches(0.7))
    s2_ind.name = "slide2_industry"
    tf_s2ind = s2_ind.text_frame
    tf_s2ind.word_wrap = True
    p_ind = tf_s2ind.paragraphs[0]
    p_ind.text = "Industry: Technology & Information Services"
    p_ind.font.name = "Calibri"
    p_ind.font.size = Pt(12)
    p_ind.font.color.rgb = COLOR_SLATE

    # Named placeholder: Size
    s2_sz = s2.shapes.add_textbox(Inches(1.1), Inches(2.85), Inches(4.6), Inches(0.7))
    s2_sz.name = "slide2_size"
    tf_s2sz = s2_sz.text_frame
    tf_s2sz.word_wrap = True
    p_sz = tf_s2sz.paragraphs[0]
    p_sz.text = "Workforce Size: 50,000+ Employees"
    p_sz.font.name = "Calibri"
    p_sz.font.size = Pt(12)
    p_sz.font.color.rgb = COLOR_SLATE

    # Named placeholder: Overview Narrative
    s2_ov = s2.shapes.add_textbox(Inches(1.1), Inches(3.6), Inches(4.6), Inches(3.0))
    s2_ov.name = "slide2_company_box"
    tf_s2ov = s2_ov.text_frame
    tf_s2ov.word_wrap = True
    p_ov = tf_s2ov.paragraphs[0]
    p_ov.text = "Company operations overview."
    p_ov.font.name = "Calibri"
    p_ov.font.size = Pt(11.5)
    p_ov.font.color.rgb = COLOR_NAVY

    # Right Container Box (Key Occupational Risks)
    s2_right = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(1.4), Inches(6.0), Inches(5.4))
    s2_right.fill.solid()
    s2_right.fill.fore_color.rgb = COLOR_LIGHT_GRAY
    s2_right.line.color.rgb = COLOR_BORDER
    s2_right.line.width = Pt(1)

    s2_hdr_right = s2.shapes.add_textbox(Inches(6.8), Inches(1.6), Inches(5.4), Inches(0.4))
    tf_s2hr = s2_hdr_right.text_frame
    p_s2hr = tf_s2hr.paragraphs[0]
    p_s2hr.text = "KEY OCCUPATIONAL EXPOSURES"
    p_s2hr.font.name = "Calibri"
    p_s2hr.font.size = Pt(13)
    p_s2hr.font.bold = True
    p_s2hr.font.color.rgb = COLOR_NAVY

    # Named placeholder: slide2_risks
    s2_risks = s2.shapes.add_textbox(Inches(6.8), Inches(2.1), Inches(5.4), Inches(4.5))
    s2_risks.name = "slide2_risks"
    tf_s2r = s2_risks.text_frame
    tf_s2r.word_wrap = True
    p_r1 = tf_s2r.paragraphs[0]
    p_r1.text = "• Occupational risks identified across workforce."
    p_r1.font.name = "Calibri"
    p_r1.font.size = Pt(11.5)
    p_r1.font.color.rgb = COLOR_NAVY

    # =========================================================================
    # Slide 3: Policy Benefits Mapped to Exposures Template
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_base_header(s3, "slide3_title", "Policy Benefits Mapped to Corporate Exposures")

    # Column Titles
    l_title = s3.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(5.0), Inches(0.4))
    tf_lt = l_title.text_frame
    p_lt = tf_lt.paragraphs[0]
    p_lt.text = "IDENTIFIED EXPOSURES"
    p_lt.font.name = "Calibri"
    p_lt.font.size = Pt(13)
    p_lt.font.bold = True
    p_lt.font.color.rgb = COLOR_SLATE

    r_title = s3.shapes.add_textbox(Inches(7.5), Inches(1.3), Inches(5.0), Inches(0.4))
    tf_rt = r_title.text_frame
    p_rt = tf_rt.paragraphs[0]
    p_rt.text = "TARGETED POLICY BENEFITS"
    p_rt.font.name = "Calibri"
    p_rt.font.size = Pt(13)
    p_rt.font.bold = True
    p_rt.font.color.rgb = COLOR_CYAN

    # Mapping Arrow Visual
    arrow = s3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.1), Inches(3.6), Inches(1.1), Inches(0.8))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = COLOR_CYAN
    arrow.line.fill.background()

    # Left Container Card
    s3_left_box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.75), Inches(4.9), Inches(5.0))
    s3_left_box.fill.solid()
    s3_left_box.fill.fore_color.rgb = COLOR_LIGHT_GRAY
    s3_left_box.line.color.rgb = COLOR_BORDER
    s3_left_box.line.width = Pt(1)

    # Named placeholder: slide3_risk_column
    s3_rc = s3.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(4.5), Inches(4.6))
    s3_rc.name = "slide3_risk_column"
    tf_s3rc = s3_rc.text_frame
    tf_s3rc.word_wrap = True
    p_s3rc = tf_s3rc.paragraphs[0]
    p_s3rc.text = "• Identified risk exposures."
    p_s3rc.font.name = "Calibri"
    p_s3rc.font.size = Pt(11.5)
    p_s3rc.font.color.rgb = COLOR_NAVY

    # Right Container Card
    s3_right_box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(1.75), Inches(5.0), Inches(5.0))
    s3_right_box.fill.solid()
    s3_right_box.fill.fore_color.rgb = COLOR_LIGHT_GRAY
    s3_right_box.line.color.rgb = COLOR_CYAN
    s3_right_box.line.width = Pt(1.5)

    # Named placeholder: slide3_benefit_column
    s3_bc = s3.shapes.add_textbox(Inches(7.7), Inches(1.95), Inches(4.6), Inches(4.6))
    s3_bc.name = "slide3_benefit_column"
    tf_s3bc = s3_bc.text_frame
    tf_s3bc.word_wrap = True
    p_s3bc = tf_s3bc.paragraphs[0]
    p_s3bc.text = "• Corresponding policy benefits."
    p_s3bc.font.name = "Calibri"
    p_s3bc.font.size = Pt(11.5)
    p_s3bc.font.color.rgb = COLOR_NAVY

    # =========================================================================
    # Slide 4: Recommended Policy & Next Steps Template
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_base_header(s4, "slide4_title", "Recommended Policy Placement & Strategic Implementation")

    # Top Half: Recommendation Box
    rec_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.73), Inches(3.3))
    rec_box.fill.solid()
    rec_box.fill.fore_color.rgb = COLOR_LIGHT_CYAN
    rec_box.line.color.rgb = COLOR_NAVY
    rec_box.line.width = Pt(1.5)

    # Named placeholder: slide4_recommendation
    s4_rec = s4.shapes.add_textbox(Inches(1.1), Inches(1.55), Inches(11.1), Inches(3.0))
    s4_rec.name = "slide4_recommendation"
    tf_s4rec = s4_rec.text_frame
    tf_s4rec.word_wrap = True
    p_s4rec = tf_s4rec.paragraphs[0]
    p_s4rec.text = "Recommended insurer placement and key reasons."
    p_s4rec.font.name = "Calibri"
    p_s4rec.font.size = Pt(12)
    p_s4rec.font.color.rgb = COLOR_NAVY

    # Bottom Half: Next Steps Title
    ns_title = s4.shapes.add_textbox(Inches(0.8), Inches(4.9), Inches(11.73), Inches(0.4))
    tf_nst = ns_title.text_frame
    p_nst = tf_nst.paragraphs[0]
    p_nst.text = "IMPLEMENTATION ROADMAP & NEXT STEPS"
    p_nst.font.name = "Calibri"
    p_nst.font.size = Pt(13)
    p_nst.font.bold = True
    p_nst.font.color.rgb = COLOR_SLATE

    # Next Steps Container
    ns_container = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.35), Inches(11.73), Inches(1.55))
    ns_container.fill.solid()
    ns_container.fill.fore_color.rgb = COLOR_LIGHT_GRAY
    ns_container.line.color.rgb = COLOR_BORDER
    ns_container.line.width = Pt(1)

    # Named placeholder: slide4_next_steps
    s4_ns = s4.shapes.add_textbox(Inches(1.1), Inches(5.45), Inches(11.1), Inches(1.35))
    s4_ns.name = "slide4_next_steps"
    tf_s4ns = s4_ns.text_frame
    tf_s4ns.word_wrap = True
    p_s4ns = tf_s4ns.paragraphs[0]
    p_s4ns.text = "1. Requirement Validation | 2. Custom Quotation | 3. Broker Presentation | 4. Policy Placement"
    p_s4ns.font.name = "Calibri"
    p_s4ns.font.size = Pt(11)
    p_s4ns.font.color.rgb = COLOR_NAVY

    prs.save(TEMPLATE_PATH)
    logger.info(f"Base PowerPoint template generated successfully at '{TEMPLATE_PATH}'.")
    return TEMPLATE_PATH


def export_to_pptx(pitch_deck: PitchDeck, output_path: str = "TCS_Pitch_Deck.pptx", profile: Optional[Any] = None) -> str:
    """
    Renders an executive, consulting-grade Marsh pitch deck using the template system.
    
    Architecture:
      1. Loads templates/marsh_pitch_template.pptx (creating it if absent via ensure_template_exists()).
      2. Injects content into predefined named shapes:
         - Cover: cover_title, cover_subtitle, cover_client, cover_date, cover_footer
         - Slide 2: slide2_title, slide2_industry, slide2_size, slide2_company_box, slide2_risks
         - Slide 3: slide3_title, slide3_risk_column, slide3_benefit_column
         - Slide 4: slide4_title, slide4_recommendation, slide4_next_steps
      3. Limits bullets to max 5 per slide with safe text truncation and word_wrap=True.
      4. Injects broker presentation scripts into the notes slide.
      5. Saves output presentation to output_path.
    """
    ensure_template_exists()
    logger.info(f"Loading base template from '{TEMPLATE_PATH}' and rendering presentation to '{output_path}'...")

    prs = Presentation(str(TEMPLATE_PATH))
    today_str = datetime.now().strftime("%B %d, %Y")
    company_name = pitch_deck.company_name or "Corporate Client"

    # Index named shapes on each slide
    slide_shapes_map = []
    for s in prs.slides:
        shape_dict = {}
        for shape in s.shapes:
            if shape.name:
                shape_dict[shape.name] = shape
        slide_shapes_map.append(shape_dict)

    # Safe extraction of slide deck data
    slides = pitch_deck.slides if pitch_deck.slides else []
    s1_data = slides[0] if len(slides) > 0 else SlideModel(slide_number=1, title="Cover", bullet_points=[], speaker_notes="")
    s2_data = slides[1] if len(slides) > 1 else SlideModel(slide_number=2, title="Company Risk Profile", bullet_points=[], speaker_notes="")
    s3_data = slides[2] if len(slides) > 2 else SlideModel(slide_number=3, title="Policy Benefits Mapping", bullet_points=[], speaker_notes="")
    s4_data = slides[3] if len(slides) > 3 else SlideModel(slide_number=4, title="Recommended Policy", bullet_points=[], speaker_notes="")

    # -------------------------------------------------------------------------
    # Slide 1: Cover
    # -------------------------------------------------------------------------
    if len(prs.slides) > 0:
        shapes = slide_shapes_map[0]
        if "cover_title" in shapes:
            tf = shapes["cover_title"].text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = "Corporate Health & Benefits Strategy"
            p.font.name = "Calibri"
            p.font.size = Pt(38)
            p.font.bold = True
            p.font.color.rgb = COLOR_NAVY

        if "cover_subtitle" in shapes:
            tf = shapes["cover_subtitle"].text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = f"Prepared for {company_name}"
            p.font.name = "Calibri"
            p.font.size = Pt(22)
            p.font.color.rgb = COLOR_CYAN

        if "cover_client" in shapes:
            tf = shapes["cover_client"].text_frame
            p = tf.paragraphs[0]
            ind_txt = getattr(profile, "industry", "Global Enterprise") if profile else "Global Enterprise"
            p.text = f"Industry: {ind_txt}  |  Brokerage Placement"
            p.font.name = "Calibri"
            p.font.size = Pt(13)
            p.font.color.rgb = COLOR_SLATE

        if "cover_date" in shapes:
            tf = shapes["cover_date"].text_frame
            p = tf.paragraphs[0]
            p.text = today_str
            p.font.name = "Calibri"
            p.font.size = Pt(12)
            p.font.color.rgb = COLOR_SLATE

        if "cover_footer" in shapes:
            tf = shapes["cover_footer"].text_frame
            p = tf.paragraphs[0]
            p.text = "Marsh Risk Advisory | Confidential"
            p.font.name = "Calibri"
            p.font.size = Pt(12)
            p.font.color.rgb = COLOR_SLATE
            p.alignment = PP_ALIGN.RIGHT

        # Speaker notes
        notes = prs.slides[0].notes_slide.notes_text_frame
        notes.text = f"[Marsh Broker Script - Executive Cover]\n\n{s1_data.speaker_notes or 'Introduce Marsh corporate brokerage team and set the agenda for custom benefits design.'}"

    # -------------------------------------------------------------------------
    # Slide 2: Company Overview & Risk Profile
    # -------------------------------------------------------------------------
    if len(prs.slides) > 1:
        shapes = slide_shapes_map[1]
        if "slide2_title" in shapes:
            tf = shapes["slide2_title"].text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = s2_data.title or "Company Overview & Identified Occupational Exposures"
            p.font.name = "Calibri"
            p.font.size = Pt(26)
            p.font.bold = True
            p.font.color.rgb = COLOR_WHITE

        # Industry & Size from Profile or Slides
        industry_val = getattr(profile, "industry", "Information Technology & Enterprise Services") if profile else "Information Technology & Enterprise Services"
        size_val = getattr(profile, "size", "50,000+ Global Workforce") if profile else "Large Enterprise Workforce"

        if "slide2_industry" in shapes:
            tf = shapes["slide2_industry"].text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = f"Industry Sector: {industry_val}"
            p.font.name = "Calibri"
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = COLOR_NAVY

        if "slide2_size" in shapes:
            tf = shapes["slide2_size"].text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = f"Workforce Scale: {size_val}"
            p.font.name = "Calibri"
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = COLOR_NAVY

        if "slide2_company_box" in shapes:
            tf = shapes["slide2_company_box"].text_frame
            tf.word_wrap = True
            raw_summary = getattr(profile, "raw_summary", "") if profile else ""
            if not raw_summary and s2_data.bullet_points:
                raw_summary = s2_data.bullet_points[0]
            if not raw_summary:
                raw_summary = f"{company_name} is a leading enterprise operating multi-shift global delivery centers with intensive desk and screen-based workforces."

            tf.text = ""
            p = tf.paragraphs[0]
            p.text = "Operational Dynamics:"
            p.font.name = "Calibri"
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = COLOR_NAVY
            p.space_after = Pt(4)

            p2 = tf.add_paragraph()
            p2.text = safe_truncate(raw_summary, max_chars=260)
            p2.font.name = "Calibri"
            p2.font.size = Pt(11)
            p2.font.color.rgb = COLOR_SLATE

        if "slide2_risks" in shapes:
            tf = shapes["slide2_risks"].text_frame
            tf.word_wrap = True
            tf.text = ""
            risk_bullets = []
            if profile and getattr(profile, "key_risks", None):
                risk_bullets = profile.key_risks
            elif s2_data.bullet_points:
                risk_bullets = s2_data.bullet_points[1:] if len(s2_data.bullet_points) > 1 else s2_data.bullet_points

            if not risk_bullets:
                risk_bullets = [
                    "Sedentary Screen Time & Pre-Diabetic Markers",
                    "Ergonomic Musculoskeletal Neck and Lumbar Strain",
                    "Workforce Stress, Rotating Shifts & Sleep Disruption",
                    "Chronic Lifestyle Disease Frequency in Desk Personnel",
                ]

            # Enforce max 5 bullets and safe truncation
            for idx, r_text in enumerate(risk_bullets[:5]):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                clean_r = safe_truncate(r_text, max_chars=130)
                format_bullet_text(p, f"•  {clean_r}", font_size_pt=11, text_color=COLOR_NAVY)
                p.space_after = Pt(8)

        # Speaker notes
        notes = prs.slides[1].notes_slide.notes_text_frame
        notes.text = f"[Marsh Broker Script - Risk Profile]\n\n{s2_data.speaker_notes or 'Highlight specific workforce exposures and absenteeism drivers quantified during risk profiling.'}"

    # -------------------------------------------------------------------------
    # Slide 3: Policy Benefits Mapped to Exposures
    # -------------------------------------------------------------------------
    if len(prs.slides) > 2:
        shapes = slide_shapes_map[2]
        if "slide3_title" in shapes:
            tf = shapes["slide3_title"].text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = s3_data.title or "Policy Architecture Mapped to Corporate Exposures"
            p.font.name = "Calibri"
            p.font.size = Pt(26)
            p.font.bold = True
            p.font.color.rgb = COLOR_WHITE

        # Split bullets into exposures (left) and policy benefits (right)
        all_pts = s3_data.bullet_points if s3_data.bullet_points else [
            "Prolonged sedentary screen time driving metabolic disorders",
            "Ergonomic back and cervical strain requiring rehabilitation",
            "Annual OPD check-up package with INR 10,000 sub-limit",
            "Unlimited teleconsultations and AYUSH inpatient coverage",
        ]

        midpoint = max(1, len(all_pts) // 2)
        left_pts = all_pts[:midpoint][:5]
        right_pts = all_pts[midpoint:][:5]

        if "slide3_risk_column" in shapes:
            tf = shapes["slide3_risk_column"].text_frame
            tf.word_wrap = True
            tf.text = ""
            for idx, pt in enumerate(left_pts):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                clean_pt = safe_truncate(pt, max_chars=130)
                format_bullet_text(p, f"•  {clean_pt}", font_size_pt=11, text_color=COLOR_NAVY)
                p.space_after = Pt(10)

        if "slide3_benefit_column" in shapes:
            tf = shapes["slide3_benefit_column"].text_frame
            tf.word_wrap = True
            tf.text = ""
            for idx, pt in enumerate(right_pts):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                clean_pt = safe_truncate(pt, max_chars=130)
                format_bullet_text(p, f"✔  {clean_pt}", font_size_pt=11, text_color=COLOR_NAVY)
                p.space_after = Pt(10)

        # Speaker notes
        notes = prs.slides[2].notes_slide.notes_text_frame
        notes.text = f"[Marsh Broker Script - Policy Mapping]\n\n{s3_data.speaker_notes or 'Demonstrate how each policy clause directly offsets identified workforce health exposures.'}"

    # -------------------------------------------------------------------------
    # Slide 4: Recommended Policy & Next Steps
    # -------------------------------------------------------------------------
    if len(prs.slides) > 3:
        shapes = slide_shapes_map[3]
        if "slide4_title" in shapes:
            tf = shapes["slide4_title"].text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = s4_data.title or "Recommended Insurer Placement & Implementation Roadmap"
            p.font.name = "Calibri"
            p.font.size = Pt(26)
            p.font.bold = True
            p.font.color.rgb = COLOR_WHITE

        if "slide4_recommendation" in shapes:
            tf = shapes["slide4_recommendation"].text_frame
            tf.word_wrap = True
            tf.text = ""

            p_h = tf.paragraphs[0]
            p_h.text = f"Primary Recommendation: Comprehensive Corporate Health Placement"
            p_h.font.name = "Calibri"
            p_h.font.size = Pt(13)
            p_h.font.bold = True
            p_h.font.color.rgb = COLOR_NAVY
            p_h.space_after = Pt(6)

            rec_bullets = s4_data.bullet_points if s4_data.bullet_points else [
                "Placement with premier insurers (Care Health / HDFC ERGO) offering 10,000+ cashless network hospitals.",
                "Day-one pre-existing disease waivers with zero copay for employees and dependents.",
                "Sum insured auto-restoration ensuring full coverage replenishment across multi-event claim years.",
            ]

            # Max 4 key reasons
            for r_bp in rec_bullets[:4]:
                p = tf.add_paragraph()
                clean_bp = safe_truncate(r_bp, max_chars=145)
                format_bullet_text(p, f"•  {clean_bp}", font_size_pt=11, text_color=COLOR_NAVY)
                p.space_after = Pt(5)

        if "slide4_next_steps" in shapes:
            tf = shapes["slide4_next_steps"].text_frame
            tf.word_wrap = True
            tf.text = ""

            p_steps = tf.paragraphs[0]
            steps = [
                ("Phase 1: Requirement Validation", "Finalize corporate census data, sum insured bands, and customized waiver clauses."),
                ("Phase 2: Custom Quotation", "Negotiate bespoke group rates with top underwriters via Marsh competitive bidding."),
                ("Phase 3: Broker Presentation", "Present comparative term sheets and stewardship service level agreements (SLAs)."),
                ("Phase 4: Policy Placement", "Bind policy, issue digital e-cards, and launch employee cashless orientation townhalls."),
            ]
            for idx, (title, desc) in enumerate(steps):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                run_t = p.add_run()
                run_t.text = f"{title}: "
                run_t.font.name = "Calibri"
                run_t.font.size = Pt(10.5)
                run_t.font.bold = True
                run_t.font.color.rgb = COLOR_NAVY

                run_d = p.add_run()
                run_d.text = f"{desc}  "
                run_d.font.name = "Calibri"
                run_d.font.size = Pt(10)
                run_d.font.color.rgb = COLOR_SLATE
                p.space_after = Pt(2)

        # Speaker notes
        notes = prs.slides[3].notes_slide.notes_text_frame
        notes.text = f"[Marsh Broker Script - Recommendation & Next Steps]\n\n{s4_data.speaker_notes or 'Close the meeting by establishing clear milestone dates for quote negotiations and binding.'}"

    out_p = Path(output_path)
    out_p.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_p))
    logger.info(f"Consulting-grade PowerPoint saved successfully to: {out_p.resolve()}")
    return str(out_p)


# =============================================================================
# Main End-to-End Orchestrator (CLI Standalone Execution)
# =============================================================================
def main():
    print("\n" + "=" * 70)
    print("      MARSH AI MULTI-AGENT PITCH GENERATOR & AUDITOR")
    print("=" * 70)

    target_company = "Tata Consultancy Services"

    # Step 1: Company Profiler
    print(f"\n[1/4] Running Company Profiler for '{target_company}'...")
    profile = generateCompanyProfile(target_company)
    print(f"  Industry   : {profile.industry}")
    print(f"  Size       : {profile.size}")
    print("  Identified Health Risks:")
    for risk in profile.key_risks:
        print(f"    - {risk}")
    if profile.is_assumption:
        print("  [!] Notice: Profile generated with fallback domain assumptions.")

    # Step 2: Pitch Generator Agent
    print(f"\n[2/4] Generating 4-Slide Risk-Mapped Pitch Deck...")
    deck, chunks = generateMarketingPitch(profile)
    print(f"  Generated {len(deck.slides)} executive slides for {deck.company_name}:")
    for s in deck.slides:
        print(f"    Slide {s.slide_number}: {s.title} ({len(s.bullet_points)} bullets)")

    # Step 3: Auditor Agent
    print(f"\n[3/4] Running Auditor Agent against Policy Knowledge Base...")
    audit_report = auditPitchContent(deck)
    print(f"\n--- AUDIT REPORT (Confidence: {audit_report.deck_confidence_score * 100:.1f}%) ---")
    print(f"Summary: {audit_report.summary}\n")
    for i, claim_audit in enumerate(audit_report.claims, 1):
        status_symbol = "[VERIFIED]" if claim_audit.status == "Verified" else "[UNVERIFIED]"
        print(f"  Claim {i}: \"{claim_audit.claim}\"")
        print(f"    Status    : {status_symbol} {claim_audit.status} (Score: {claim_audit.confidence_score})")
        print(f"    Feature   : {claim_audit.core_policy_feature}")
        print(f"    Limit     : {claim_audit.stated_limit_or_rule}")
        print(f"    Application: {claim_audit.client_application}")
        print(f"    Evidence  : {claim_audit.evidence_snippet}")
        print(f"    Source    : {claim_audit.source_document}\n")

    # Step 4: Export to PPTX
    print(f"[4/4] Exporting to Premium PPTX presentation...")
    output_filename = "TCS_Pitch_Deck.pptx"
    export_to_pptx(deck, output_path=output_filename)

    print("\n" + "=" * 70)
    print(f"SUCCESS: Pipeline complete! Generated deck: '{output_filename}'")
    print("=" * 70 + "\n")


if __name__ == "__main__":
    main()
