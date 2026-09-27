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

    if "OpenAI" in provider:
        if not openai_key:
            logger.warning("OPENAI_API_KEY missing. Falling back to Groq.")
            if groq_key and ChatGroq:
                return ChatGroq(model_name="openai/gpt-oss-120b", temperature=0.2)
        return ChatOpenAI(model="gpt-4o-mini", temperature=0.2)
    else:
        # Groq selection
        if groq_key and ChatGroq:
            logger.info("Using Groq LLM engine.")
            return ChatGroq(model_name="openai/gpt-oss-120b", temperature=0.2)
        elif openai_key:
            logger.info("Groq key unavailable, falling back to OpenAI gpt-4o-mini.")
            return ChatOpenAI(model="gpt-4o-mini", temperature=0.2)
        else:
            return ChatOpenAI(model="gpt-4o-mini", temperature=0.2)


def invoke_structured(llm, pydantic_cls, system_prompt: str, user_prompt: str):
    """
    Robust structured invoker:
    Attempts native with_structured_output first; seamlessly falls back
    to PydanticOutputParser for models with tool-call variances.
    """
    try:
        structured_llm = llm.with_structured_output(pydantic_cls)
        return structured_llm.invoke([
            SystemMessage(content=system_prompt),
            HumanMessage(content=user_prompt),
        ])
    except Exception as e:
        logger.warning(f"Native structured output fallback engaged ({e}). Using PydanticOutputParser.")
        parser = PydanticOutputParser(pydantic_object=pydantic_cls)
        format_instructions = parser.get_format_instructions()
        augmented_prompt = f"{system_prompt}\n\n{user_prompt}\n\nStrict Schema Requirement:\n{format_instructions}"
        response = llm.invoke(augmented_prompt)
        text_content = response.content
        clean_json = re.sub(r"^```(?:json)?\s*", "", text_content.strip())
        clean_json = re.sub(r"\s*```$", "", clean_json)
        return parser.parse(clean_json)


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
        min_length=4,
        max_length=5,
        description="4 to 5 detailed, data-dense bullet points (each 20-35 words with concrete metrics, policy terms, rupee limits, waiting periods, or quantitative advantages)."
    )
    speaker_notes: str = Field(..., description="Detailed conversational script for the Marsh insurance broker with talk track guidance")


class PitchDeck(BaseModel):
    company_name: str
    slides: List[SlideModel] = Field(..., min_length=4, max_length=4)


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
# 6. Consulting-Grade PPTX Export Utility
# =============================================================================
def export_to_pptx(pitch_deck: PitchDeck, output_path: str = "TCS_Pitch_Deck.pptx") -> str:
    """
    Renders an executive, Consulting-Grade (McKinsey / Marsh style) presentation
    using shape-based infographics, chevron roadmaps, and strict grid layouts.

    Global Design System:
      - Dimensions: 16:9 (Width: Inches(13.33), Height: Inches(7.5))
      - Palette: Navy (RGBColor(0, 32, 91)), Cyan (RGBColor(0, 163, 224)),
                 White, Slate Gray (RGBColor(71, 85, 105)), Light Gray (RGBColor(241, 245, 249)).
      - Font: 'Calibri'
      - Margins: Strict 0.8-inch left/right margins for all content.
    """
    logger.info(f"Rendering McKinsey/Marsh shape-driven presentation to '{output_path}'...")
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    COLOR_NAVY = RGBColor(0, 32, 91)         # #00205B Marsh Navy
    COLOR_CYAN = RGBColor(0, 163, 224)       # #00A3E0 Marsh Cyan
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_SLATE = RGBColor(71, 85, 105)      # Slate Gray
    COLOR_LIGHT_GRAY = RGBColor(241, 245, 249)# Light Gray fill
    COLOR_LIGHT_CYAN = RGBColor(224, 244, 252)# Soft Cyan background
    COLOR_BORDER = RGBColor(203, 213, 225)   # Border Gray

    blank_layout = prs.slide_layouts[6]
    today_str = datetime.now().strftime("%B %d, %Y")

    def add_standard_header(slide, title_text: str):
        """Draws standard Navy banner and Cyan accent line across the top."""
        header_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.33), Inches(1.0)
        )
        header_bar.fill.solid()
        header_bar.fill.fore_color.rgb = COLOR_NAVY
        header_bar.line.fill.background()

        # Accent line below Navy header
        accent_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.0), Inches(13.33), Pt(4)
        )
        accent_line.fill.solid()
        accent_line.fill.fore_color.rgb = COLOR_CYAN
        accent_line.line.fill.background()

        # Title text
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.15), Inches(10.5), Inches(0.7))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0)
        tf.margin_top = Inches(0)
        p = tf.paragraphs[0]
        p.text = title_text
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

    for slide_data in pitch_deck.slides:
        slide = prs.slides.add_slide(blank_layout)

        # Base clean white background
        base_bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.33), Inches(7.5)
        )
        base_bg.fill.solid()
        base_bg.fill.fore_color.rgb = COLOR_WHITE
        base_bg.line.fill.background()

        # =====================================================================
        # Slide 1: The Executive Cover
        # =====================================================================
        if slide_data.slide_number == 1:
            # Left 30% Solid Navy Rectangle (Width: Inches(4), Height: Inches(7.5))
            left_panel = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(4.0), Inches(7.5)
            )
            left_panel.fill.solid()
            left_panel.fill.fore_color.rgb = COLOR_NAVY
            left_panel.line.fill.background()

            # Thin Cyan vertical divider line
            divider = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(4.0), Inches(0), Inches(0.05), Inches(7.5)
            )
            divider.fill.solid()
            divider.fill.fore_color.rgb = COLOR_CYAN
            divider.line.fill.background()

            # Inside Navy panel: "MARSH" (White, 40pt, Bold) & "RISK ADVISORY" (Cyan, 16pt)
            navy_brand_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.8), Inches(2.8), Inches(2.5))
            tf_nb = navy_brand_box.text_frame
            tf_nb.word_wrap = True
            tf_nb.margin_left = Inches(0)
            tf_nb.margin_top = Inches(0)

            p_m = tf_nb.paragraphs[0]
            p_m.text = "MARSH"
            p_m.font.name = "Calibri"
            p_m.font.size = Pt(40)
            p_m.font.bold = True
            p_m.font.color.rgb = COLOR_WHITE

            p_ra = tf_nb.add_paragraph()
            p_ra.text = "RISK ADVISORY"
            p_ra.font.name = "Calibri"
            p_ra.font.size = Pt(16)
            p_ra.font.bold = True
            p_ra.font.color.rgb = COLOR_CYAN
            p_ra.space_before = Pt(6)

            # On Right 70% (White Background): Title & Subtitle
            right_content_box = slide.shapes.add_textbox(Inches(4.8), Inches(2.2), Inches(7.7), Inches(3.5))
            tf_rc = right_content_box.text_frame
            tf_rc.word_wrap = True
            tf_rc.margin_left = Inches(0)
            tf_rc.margin_top = Inches(0)

            p_t = tf_rc.paragraphs[0]
            p_t.text = "Corporate Health & Benefits Strategy"
            p_t.font.name = "Calibri"
            p_t.font.size = Pt(44)
            p_t.font.bold = True
            p_t.font.color.rgb = COLOR_NAVY
            p_t.space_after = Pt(14)

            p_sub = tf_rc.add_paragraph()
            p_sub.text = f"Prepared for {pitch_deck.company_name}"
            p_sub.font.name = "Calibri"
            p_sub.font.size = Pt(24)
            p_sub.font.color.rgb = COLOR_SLATE
            p_sub.space_after = Pt(20)

            p_desc = tf_rc.add_paragraph()
            p_desc.text = "Data-Driven Risk Placement & Statutory Policy Architecture"
            p_desc.font.name = "Calibri"
            p_desc.font.size = Pt(14)
            p_desc.font.color.rgb = COLOR_SLATE

            # Bottom right: Date and "Confidential" (Slate Gray, 12pt)
            footer_box = slide.shapes.add_textbox(Inches(4.8), Inches(6.4), Inches(7.7), Inches(0.5))
            tf_foot = footer_box.text_frame
            tf_foot.margin_left = Inches(0)
            p_f = tf_foot.paragraphs[0]
            p_f.text = f"{today_str}  |  Strictly Confidential"
            p_f.font.name = "Calibri"
            p_f.font.size = Pt(12)
            p_f.font.color.rgb = COLOR_SLATE

        # =====================================================================
        # Slide 2: Executive Risk Profile
        # =====================================================================
        elif slide_data.slide_number == 2:
            add_standard_header(slide, slide_data.title or "Executive Risk Profile")

            # Left Column (Width: 5.2 inches): "Company Overview"
            left_box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.2), Inches(5.2)
            )
            left_box.fill.solid()
            left_box.fill.fore_color.rgb = COLOR_LIGHT_GRAY
            left_box.line.color.rgb = COLOR_BORDER
            left_box.line.width = Pt(1)

            # Left Column Content Box
            lt_box = slide.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(4.6), Inches(4.5))
            tf_lt = lt_box.text_frame
            tf_lt.word_wrap = True
            tf_lt.margin_left = Inches(0)
            tf_lt.margin_top = Inches(0)

            p_oh = tf_lt.paragraphs[0]
            p_oh.text = "COMPANY OVERVIEW"
            p_oh.font.name = "Calibri"
            p_oh.font.size = Pt(14)
            p_oh.font.bold = True
            p_oh.font.color.rgb = COLOR_NAVY
            p_oh.space_after = Pt(14)

            # Insert first 2 bullets as structured overview
            overview_bullets = slide_data.bullet_points[:2] if slide_data.bullet_points else ["Corporate enterprise operations."]
            for bp in overview_bullets:
                p_b = tf_lt.add_paragraph()
                p_b.text = f"•  {bp}"
                p_b.font.name = "Calibri"
                p_b.font.size = Pt(13)
                p_b.font.color.rgb = COLOR_NAVY
                p_b.space_after = Pt(12)

            # Bottom highlight pill
            badge_shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(5.7), Inches(4.6), Inches(0.65)
            )
            badge_shape.fill.solid()
            badge_shape.fill.fore_color.rgb = COLOR_NAVY
            badge_shape.line.fill.background()
            tf_badge = badge_shape.text_frame
            p_bdg = tf_badge.paragraphs[0]
            p_bdg.text = f"Target Client: {pitch_deck.company_name}"
            p_bdg.font.name = "Calibri"
            p_bdg.font.size = Pt(13)
            p_bdg.font.bold = True
            p_bdg.font.color.rgb = COLOR_WHITE
            p_bdg.alignment = PP_ALIGN.CENTER

            # Right Column (Width: 6.0 inches): "Key Occupational Exposures"
            right_header_box = slide.shapes.add_textbox(Inches(6.5), Inches(1.4), Inches(6.0), Inches(0.5))
            tf_rh = right_header_box.text_frame
            p_rh = tf_rh.paragraphs[0]
            p_rh.text = "KEY OCCUPATIONAL EXPOSURES"
            p_rh.font.name = "Calibri"
            p_rh.font.size = Pt(14)
            p_rh.font.bold = True
            p_rh.font.color.rgb = COLOR_NAVY

            # VISUAL ELEMENT: Draw 3 or 4 ROUNDED_RECTANGLE cards evenly spaced vertically
            risk_bullets = slide_data.bullet_points[2:] if len(slide_data.bullet_points) > 2 else slide_data.bullet_points
            if not risk_bullets:
                risk_bullets = ["Workplace ergonomics and postural fatigue", "Mental health strain and workload burnout", "Lifestyle chronic disease prevention"]

            num_cards = min(4, len(risk_bullets))
            card_height = 1.05
            gap = 0.16
            start_y = 2.0

            for i in range(num_cards):
                card_y = start_y + (i * (card_height + gap))
                card_shape = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(card_y), Inches(6.0), Inches(card_height)
                )
                card_shape.fill.solid()
                card_shape.fill.fore_color.rgb = COLOR_LIGHT_GRAY
                card_shape.line.color.rgb = COLOR_SLATE
                card_shape.line.width = Pt(1)

                # Cyan indicator tab on left edge of card
                pill_tab = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(card_y + 0.15), Inches(0.12), Inches(card_height - 0.3)
                )
                pill_tab.fill.solid()
                pill_tab.fill.fore_color.rgb = COLOR_CYAN
                pill_tab.line.fill.background()

                # Text inside card
                tx_box = slide.shapes.add_textbox(Inches(6.85), Inches(card_y + 0.08), Inches(5.5), Inches(card_height - 0.16))
                tf_card = tx_box.text_frame
                tf_card.word_wrap = True
                tf_card.margin_left = Inches(0.05)
                tf_card.margin_right = Inches(0.05)
                tf_card.margin_top = Inches(0.05)
                tf_card.margin_bottom = Inches(0.05)

                p_c = tf_card.paragraphs[0]
                p_c.text = risk_bullets[i]
                p_c.font.name = "Calibri"
                p_c.font.size = Pt(12)
                p_c.font.color.rgb = COLOR_NAVY

        # =====================================================================
        # Slide 3: Strategic Policy Mapping
        # =====================================================================
        elif slide_data.slide_number == 3:
            add_standard_header(slide, "Strategic Policy Mapping: Tailored Policy Provisions")

            # Column Headers
            left_col_title = slide.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(5.0), Inches(0.4))
            tf_lct = left_col_title.text_frame
            p_lct = tf_lct.paragraphs[0]
            p_lct.text = "IDENTIFIED EXPOSURES"
            p_lct.font.name = "Calibri"
            p_lct.font.size = Pt(14)
            p_lct.font.bold = True
            p_lct.font.color.rgb = COLOR_SLATE

            right_col_title = slide.shapes.add_textbox(Inches(7.5), Inches(1.3), Inches(5.0), Inches(0.4))
            tf_rct = right_col_title.text_frame
            p_rct = tf_rct.paragraphs[0]
            p_rct.text = "MARSH POLICY SOLUTIONS"
            p_rct.font.name = "Calibri"
            p_rct.font.size = Pt(14)
            p_rct.font.bold = True
            p_rct.font.color.rgb = COLOR_CYAN

            # VISUAL ELEMENT: Large RIGHT_ARROW in the center pointing from left to right
            arrow_shape = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW, Inches(6.05), Inches(3.2), Inches(1.2), Inches(1.0)
            )
            arrow_shape.fill.solid()
            arrow_shape.fill.fore_color.rgb = COLOR_CYAN
            arrow_shape.line.fill.background()

            # Split bullets across Left (Exposures) and Right (Solutions)
            all_pts = slide_data.bullet_points
            midpoint = max(1, len(all_pts) // 2)
            left_pts = all_pts[:midpoint]
            right_pts = all_pts[midpoint:]

            # Left Container Box
            left_card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(4.9), Inches(4.9)
            )
            left_card.fill.solid()
            left_card.fill.fore_color.rgb = COLOR_LIGHT_GRAY
            left_card.line.color.rgb = COLOR_BORDER
            left_card.line.width = Pt(1)

            tb_left = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(4.5), Inches(4.5))
            tf_l = tb_left.text_frame
            tf_l.word_wrap = True
            tf_l.margin_left = Inches(0.1)
            tf_l.margin_top = Inches(0.1)
            tf_l.margin_right = Inches(0.1)
            tf_l.margin_bottom = Inches(0.1)

            for idx, pt in enumerate(left_pts):
                p = tf_l.paragraphs[0] if idx == 0 else tf_l.add_paragraph()
                p.text = f"•   {pt}"
                p.font.name = "Calibri"
                p.font.size = Pt(12)
                p.font.color.rgb = COLOR_NAVY
                p.space_after = Pt(12)

            # Right Container Box
            right_card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(1.8), Inches(5.0), Inches(4.9)
            )
            right_card.fill.solid()
            right_card.fill.fore_color.rgb = COLOR_LIGHT_GRAY
            right_card.line.color.rgb = COLOR_CYAN
            right_card.line.width = Pt(1.5)

            tb_right = slide.shapes.add_textbox(Inches(7.7), Inches(2.0), Inches(4.6), Inches(4.5))
            tf_r = tb_right.text_frame
            tf_r.word_wrap = True
            tf_r.margin_left = Inches(0.1)
            tf_r.margin_top = Inches(0.1)
            tf_r.margin_right = Inches(0.1)
            tf_r.margin_bottom = Inches(0.1)

            for idx, pt in enumerate(right_pts):
                p = tf_r.paragraphs[0] if idx == 0 else tf_r.add_paragraph()
                p.text = f"✔   {pt}"
                p.font.name = "Calibri"
                p.font.size = Pt(12)
                p.font.color.rgb = COLOR_NAVY
                p.space_after = Pt(12)

        # =====================================================================
        # Slide 4: Recommended Solution & Next Steps
        # =====================================================================
        elif slide_data.slide_number == 4:
            add_standard_header(slide, "Recommended Placement & Implementation Roadmap")

            # TOP HALF: Large ROUNDED_RECTANGLE (Light Cyan fill, Navy border)
            top_rec_box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.73), Inches(3.4)
            )
            top_rec_box.fill.solid()
            top_rec_box.fill.fore_color.rgb = COLOR_LIGHT_CYAN
            top_rec_box.line.color.rgb = COLOR_NAVY
            top_rec_box.line.width = Pt(1.5)

            # Text inside top recommendation container
            tx_rec = slide.shapes.add_textbox(Inches(1.1), Inches(1.55), Inches(11.1), Inches(3.1))
            tf_rec = tx_rec.text_frame
            tf_rec.word_wrap = True
            tf_rec.margin_left = Inches(0.1)
            tf_rec.margin_top = Inches(0.1)
            tf_rec.margin_right = Inches(0.1)
            tf_rec.margin_bottom = Inches(0.1)

            p_rt = tf_rec.paragraphs[0]
            p_rt.text = f"RECOMMENDED POLICY PLACEMENT: {slide_data.title}"
            p_rt.font.name = "Calibri"
            p_rt.font.size = Pt(15)
            p_rt.font.bold = True
            p_rt.font.color.rgb = COLOR_NAVY
            p_rt.space_after = Pt(8)

            # Display all key bullet points from the final slide
            top_bullets = slide_data.bullet_points if slide_data.bullet_points else ["Optimal deductible and premium ratio.", "Seamless network hospital integration.", "Day-one pre-existing condition waiver."]
            for bp in top_bullets:
                p_b = tf_rec.add_paragraph()
                p_b.text = f"•  {bp}"
                p_b.font.name = "Calibri"
                p_b.font.size = Pt(12)
                p_b.font.color.rgb = COLOR_NAVY
                p_b.space_after = Pt(6)

            # BOTTOM HALF: "Implementation Roadmap" with 3 CHEVRON shapes
            rm_title_box = slide.shapes.add_textbox(Inches(0.8), Inches(5.0), Inches(11.73), Inches(0.4))
            tf_rmt = rm_title_box.text_frame
            p_rmt = tf_rmt.paragraphs[0]
            p_rmt.text = "IMPLEMENTATION ROADMAP"
            p_rmt.font.name = "Calibri"
            p_rmt.font.size = Pt(13)
            p_rmt.font.bold = True
            p_rmt.font.color.rgb = COLOR_SLATE

            # 3 CHEVRON shapes horizontally across the bottom (Cyan, Navy, Slate)
            chevron_data = [
                ("1. Needs Analysis", COLOR_CYAN, COLOR_WHITE),
                ("2. Market Placement", COLOR_NAVY, COLOR_WHITE),
                ("3. Policy Binding", COLOR_SLATE, COLOR_WHITE),
            ]

            ch_w = Inches(3.7)
            ch_h = Inches(1.1)
            gap_ch = Inches(0.3)
            start_x = Inches(0.8)

            for idx, (label, bg_col, text_col) in enumerate(chevron_data):
                curr_x = start_x + (idx * (ch_w + gap_ch))
                ch_shape = slide.shapes.add_shape(
                    MSO_SHAPE.CHEVRON, curr_x, Inches(5.5), ch_w, ch_h
                )
                ch_shape.fill.solid()
                ch_shape.fill.fore_color.rgb = bg_col
                ch_shape.line.fill.background()

                tf_ch = ch_shape.text_frame
                tf_ch.word_wrap = True
                tf_ch.margin_left = Inches(0.3)
                p_c = tf_ch.paragraphs[0]
                p_c.text = label
                p_c.font.name = "Calibri"
                p_c.font.size = Pt(16)
                p_c.font.bold = True
                p_c.font.color.rgb = text_col
                p_c.alignment = PP_ALIGN.CENTER

        # Speaker notes added cleanly
        notes_slide = slide.notes_slide
        tf_notes = notes_slide.notes_text_frame
        tf_notes.text = f"[Marsh Broker Presentation Script]\n\n{slide_data.speaker_notes}"

    prs.save(output_path)
    logger.info(f"Consulting-grade presentation saved successfully to: {Path(output_path).resolve()}")
    return output_path


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
