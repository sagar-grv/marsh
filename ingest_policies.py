"""
ingest_policies.py - Data Ingestion and Vector Store Pipeline
=============================================================
Extracts text from insurance policy brochures, cleans formatting/tables,
chunks text semantically with LangChain, generates embeddings using HuggingFace,
and stores the results in a persistent ChromaDB vector store.
"""

import os
import re
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional

import pymupdf  # PyMuPDF
from langchain_core.documents import Document
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_chroma import Chroma

# --- Logging Configuration ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("PolicyIngestion")

# --- Configuration Constants ---
TARGET_PDFS = [
    "HDFC Product Brochure.pdf",
    "Niva Bupa Product Brochure.pdf",
    "Care Health Product Brochure.pdf",
    "ABHI Product Brochure.pdf",
]

# Possible locations where PDFs might reside (current directory or Policy Documents/)
SEARCH_DIRECTORIES = [
    Path("Policy Documents"),
    Path("."),
]

VECTOR_DB_DIR = "marsh_policy_db"
COLLECTION_NAME = "marsh_policies"
EMBEDDING_MODEL_NAME = "BAAI/bge-small-en-v1.5"
CHUNK_SIZE = 512
CHUNK_OVERLAP = 100


def locate_pdf(pdf_name: str) -> Optional[Path]:
    """Search for a PDF across defined search paths."""
    for search_dir in SEARCH_DIRECTORIES:
        candidate_path = search_dir / pdf_name
        if candidate_path.is_file():
            return candidate_path
    return None


def clean_text(raw_text: str) -> str:
    """
    Cleans raw text extracted from PDFs:
    - Normalizes non-breaking spaces and irregular whitespace.
    - Joins words broken across line breaks with hyphens (e.g. 'inves-\ntigation').
    - Collapses excessive whitespace within lines (table columns/tabs).
    - Reduces repeated blank lines to double newlines.
    """
    if not raw_text:
        return ""

    # Replace non-breaking spaces and tabs with standard space
    text = raw_text.replace("\u00a0", " ").replace("\t", " ")

    # Reconnect words split across lines by hyphens
    text = re.sub(r"(\b\w+)-\n(\w+\b)", r"\1\2", text)

    # Replace runs of spaces with a single space while keeping line structures
    lines = [re.sub(r" {2,}", " ", line).strip() for line in text.splitlines()]

    # Reconstruct text and collapse multiple empty lines into at most two newlines
    cleaned = "\n".join(lines)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)

    return cleaned.strip()


def extract_text_from_pdf(file_path: Path) -> List[Dict[str, Any]]:
    """
    Extracts and cleans page-level text from a PDF file using PyMuPDF (fitz).
    Returns a list of dicts with page text and page metadata.
    """
    extracted_pages = []
    try:
        doc = pymupdf.open(file_path)
        logger.info(f"Extracting '{file_path.name}' ({len(doc)} pages)...")

        for page_idx, page in enumerate(doc):
            raw_text = page.get_text("text")
            cleaned = clean_text(raw_text)

            # Check if page has low text density but contains infographic/table images
            images = page.get_images(full=True)
            if len(cleaned) < 80 and len(images) > 0:
                logger.info(f"  [Graphic Notice] Page {page_idx + 1} has sparse text ({len(cleaned)} chars) and {len(images)} images.")

            # Skip empty pages
            if not cleaned:
                continue

            extracted_pages.append({
                "page_number": page_idx + 1,
                "text": cleaned,
            })

        doc.close()
    except Exception as e:
        logger.error(f"Failed to read '{file_path}': {e}", exc_info=True)

    return extracted_pages


def build_documents(extracted_pages: List[Dict[str, Any]], pdf_name: str) -> List[Document]:
    """Converts extracted page dictionaries into LangChain Document instances."""
    documents = []
    for item in extracted_pages:
        doc = Document(
            page_content=item["text"],
            metadata={
                "source_document": pdf_name,
                "page": item["page_number"],
            },
        )
        documents.append(doc)
    return documents


def main():
    logger.info("=" * 60)
    logger.info("Starting Marsh Health Insurance Policy Ingestion Pipeline")
    logger.info("=" * 60)

    # 1. Locate and Extract PDFs
    all_raw_documents: List[Document] = []
    found_count = 0

    for pdf_name in TARGET_PDFS:
        pdf_path = locate_pdf(pdf_name)
        if not pdf_path:
            logger.warning(f"⚠️  Missing PDF: '{pdf_name}' could not be found. Skipping.")
            continue

        found_count += 1
        pages_data = extract_text_from_pdf(pdf_path)
        docs = build_documents(pages_data, pdf_name=pdf_name)
        all_raw_documents.extend(docs)
        logger.info(f"Loaded {len(docs)} valid pages from '{pdf_name}'.")

    if not all_raw_documents:
        logger.error("❌ No text extracted from any PDF documents. Aborting ingestion.")
        return

    logger.info(
        f"Total documents extracted: {len(all_raw_documents)} pages across {found_count}/{len(TARGET_PDFS)} brochures."
    )

    # 2. Text Splitting into Semantic Chunks
    logger.info(f"Splitting text with chunk_size={CHUNK_SIZE}, chunk_overlap={CHUNK_OVERLAP}...")
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=CHUNK_SIZE,
        chunk_overlap=CHUNK_OVERLAP,
        length_function=len,
        separators=["\n\n", "\n", ". ", " ", ""],
    )

    chunks = splitter.split_documents(all_raw_documents)
    logger.info(f"Created {len(chunks)} text chunks.")

    # 3. Initialize HuggingFace Embeddings
    logger.info(f"Initializing embedding model: '{EMBEDDING_MODEL_NAME}'...")
    embeddings = HuggingFaceEmbeddings(
        model_name=EMBEDDING_MODEL_NAME,
        model_kwargs={"device": "cpu"},
        encode_kwargs={"normalize_embeddings": True},
    )

    # 4. Ingest and Persist in ChromaDB
    logger.info("=" * 60)
    logger.info("IMPORTANT: If changing embedding models, ensure the previous 'marsh_policy_db'")
    logger.info("directory was deleted/overwritten to prevent mixed vector dimensions.")
    logger.info(f"Indexing and storing chunks with '{EMBEDDING_MODEL_NAME}' in ChromaDB at '{VECTOR_DB_DIR}'...")
    logger.info("=" * 60)
    vector_db = Chroma.from_documents(
        documents=chunks,
        embedding=embeddings,
        collection_name=COLLECTION_NAME,
        persist_directory=VECTOR_DB_DIR,
    )

    logger.info("=" * 60)
    logger.info(f"SUCCESS: Ingestion complete!")
    logger.info(f"  - Database directory: '{VECTOR_DB_DIR}'")
    logger.info(f"  - Collection name: '{COLLECTION_NAME}'")
    logger.info(f"  - Total chunks indexed: {len(chunks)}")
    logger.info("=" * 60)

    # Quick verification query (BGE models use instruction prefix for queries)
    raw_query = "What is the pre-existing disease waiting period?"
    query = f"Represent this sentence for searching relevant passages: {raw_query}"
    results = vector_db.similarity_search(query, k=2)
    logger.info(f"\n[Verification Query] '{raw_query}'")
    for idx, doc in enumerate(results, 1):
        source = doc.metadata.get("source_document", "Unknown")
        page = doc.metadata.get("page", "?")
        snippet = doc.page_content[:140].replace("\n", " ")
        logger.info(f"  Result {idx} ({source}, pg {page}): {snippet}...")


if __name__ == "__main__":
    main()
