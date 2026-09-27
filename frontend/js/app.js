/**
 * frontend/js/app.js - Enterprise Client Logic (Anti-Slop Corporate Edition)
 */

// Global State
let currentPitchData = null;
let activeTab = 'summary';

// DOM Elements
const documentsListEl = document.getElementById('documents-list');
const docCounterEl = document.getElementById('doc-counter');
const generatorForm = document.getElementById('generator-form');
const generateBtn = document.getElementById('generate-btn');
const btnText = document.getElementById('btn-text');
const btnSpinner = document.getElementById('btn-spinner');

const emptyState = document.getElementById('empty-state');
const loadingState = document.getElementById('loading-state');
const resultsContainer = document.getElementById('results-container');
const errorBanner = document.getElementById('error-banner');
const errorMessageEl = document.getElementById('error-message');

document.addEventListener('DOMContentLoaded', () => {
  fetchDocuments();
  setupEventListeners();
});

function setupEventListeners() {
  generatorForm.addEventListener('submit', handleFormSubmit);
}

// =============================================================================
// 1. Fetch Available Baseline Documents
// =============================================================================
async function fetchDocuments() {
  try {
    const res = await fetch('/api/documents');
    if (!res.ok) throw new Error('Failed to retrieve baseline documents.');
    const data = await res.json();
    renderDocumentCheckboxes(data.documents);
  } catch (err) {
    console.error('Error fetching documents:', err);
    documentsListEl.innerHTML = `
      <div class="text-xs text-rose-600 py-1 font-mono">Failed to load policy brochures.</div>
    `;
  }
}

function renderDocumentCheckboxes(docs) {
  if (!docs || docs.length === 0) {
    documentsListEl.innerHTML = '<div class="text-xs text-slate-400 font-mono">No documents found.</div>';
    return;
  }

  documentsListEl.innerHTML = '';
  docs.forEach((doc, idx) => {
    const div = document.createElement('div');
    div.className = 'flex items-center space-x-2 py-0.5';
    div.innerHTML = `
      <input 
        type="checkbox" 
        id="doc-${idx}" 
        value="${doc.filename}" 
        checked 
        class="doc-checkbox rounded-sm text-[#00205B] focus:ring-0 focus:ring-offset-0 h-3.5 w-3.5 border-slate-300 cursor-pointer"
        onchange="updateDocCounter()"
      />
      <label for="doc-${idx}" class="text-xs text-slate-700 cursor-pointer font-medium select-none flex-1 truncate">
        ${doc.display_name}
      </label>
    `;
    documentsListEl.appendChild(div);
  });

  updateDocCounter();
}

function updateDocCounter() {
  const checkboxes = document.querySelectorAll('.doc-checkbox');
  const checked = document.querySelectorAll('.doc-checkbox:checked');
  docCounterEl.textContent = `${checked.length} / ${checkboxes.length} Active`;

  if (checked.length === 0) {
    generateBtn.disabled = true;
    docCounterEl.classList.add('text-rose-600');
  } else {
    generateBtn.disabled = false;
    docCounterEl.classList.remove('text-rose-600');
  }
}

// =============================================================================
// 2. Submit & Execution
// =============================================================================
async function handleFormSubmit(e) {
  e.preventDefault();
  dismissError();

  const companyName = document.getElementById('company-name').value.trim();
  const llmProvider = document.getElementById('llm-provider').value;
  const checkedDocs = Array.from(document.querySelectorAll('.doc-checkbox:checked')).map(cb => cb.value);

  if (!companyName) {
    showError('Please enter a target company name.');
    return;
  }
  if (checkedDocs.length === 0) {
    showError('Please select at least one baseline policy document.');
    return;
  }

  setLoading(true);

  try {
    const res = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        company_name: companyName,
        selected_documents: checkedDocs,
        llm_provider: llmProvider,
      }),
    });

    if (!res.ok) {
      const errData = await res.json().catch(() => ({ detail: 'Server error encountered.' }));
      throw new Error(errData.detail || 'Failed to generate presentation and audit.');
    }

    const data = await res.json();
    currentPitchData = data;
    renderResults(data);

  } catch (err) {
    console.error('Generation Error:', err);
    showError(err.message || 'An error occurred during multi-agent execution.');
    emptyState.classList.remove('hidden');
    resultsContainer.classList.add('hidden');
  } finally {
    setLoading(false);
  }
}

function setLoading(isLoading) {
  if (isLoading) {
    generateBtn.disabled = true;
    btnText.textContent = 'Executing Pipeline...';
    btnSpinner.classList.remove('hidden');
    emptyState.classList.add('hidden');
    resultsContainer.classList.add('hidden');
    loadingState.classList.remove('hidden');
  } else {
    generateBtn.disabled = false;
    btnText.textContent = 'Run Multi-Agent Placement';
    btnSpinner.classList.add('hidden');
    loadingState.classList.add('hidden');
  }
}

// =============================================================================
// 3. Render Results
// =============================================================================
function renderResults(data) {
  emptyState.classList.add('hidden');
  loadingState.classList.add('hidden');
  resultsContainer.classList.remove('hidden');

  const { profile, pitch_deck, audit_report, retrieved_chunks, downloads } = data;

  // 1. KPI Cards
  renderKPICards(audit_report);

  // 2. Executive Summary Tab
  renderExecutiveSummary(profile, audit_report);

  // 3. Pitch Deck Tab
  renderPitchDeck(pitch_deck);

  // 4. Compliance Audit Matrix (Table)
  renderComplianceAudit(audit_report);

  // 5. RAG Grounding Evidence Tab
  renderRAGEvidence(retrieved_chunks);

  // 6. Downloads Hub (With PDF prominence)
  renderDownloadsHub(downloads);

  // Default tab
  switchTab('summary');
}

function renderKPICards(audit) {
  const confPct = Math.round(audit.deck_confidence_score * 100);
  const totalClaims = audit.claims ? audit.claims.length : 0;
  let verifiedCount = 0;
  let flaggedCount = 0;

  (audit.claims || []).forEach(c => {
    if (c.status.toLowerCase() === 'verified') {
      verifiedCount++;
    } else {
      flaggedCount++;
    }
  });

  const kpiConfidence = document.getElementById('kpi-confidence');
  const kpiConfidenceStatus = document.getElementById('kpi-confidence-status');
  const kpiProgress = document.getElementById('kpi-progress');
  const kpiTotalClaims = document.getElementById('kpi-total-claims');
  const kpiVerifiedClaims = document.getElementById('kpi-verified-claims');
  const kpiFlaggedClaims = document.getElementById('kpi-flagged-claims');

  kpiConfidence.textContent = `${confPct}%`;
  kpiProgress.style.width = `${confPct}%`;

  if (confPct >= 70) {
    kpiConfidenceStatus.textContent = 'PASS';
    kpiConfidenceStatus.className = 'text-[10px] font-bold px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200 uppercase';
    kpiProgress.className = 'bg-emerald-600 h-1';
  } else {
    kpiConfidenceStatus.textContent = 'REVIEW';
    kpiConfidenceStatus.className = 'text-[10px] font-bold px-1.5 py-0.5 rounded bg-rose-50 text-rose-700 border border-rose-200 uppercase';
    kpiProgress.className = 'bg-rose-600 h-1';
  }

  kpiTotalClaims.textContent = totalClaims;
  kpiVerifiedClaims.textContent = verifiedCount;
  kpiFlaggedClaims.textContent = flaggedCount;
}

function renderExecutiveSummary(profile, audit) {
  document.getElementById('summary-company-title').textContent = profile.company_name;
  document.getElementById('summary-industry').textContent = profile.industry || 'Information Technology';
  document.getElementById('summary-size').textContent = profile.size || 'Enterprise';
  document.getElementById('summary-audit-text').textContent = audit.summary;

  const risksContainer = document.getElementById('summary-risks');
  risksContainer.innerHTML = '';
  (profile.key_risks || []).forEach(risk => {
    const badge = document.createElement('span');
    badge.className = 'px-2.5 py-1 rounded-sm text-xs font-medium bg-white border border-slate-300 text-slate-800';
    badge.textContent = risk;
    risksContainer.appendChild(badge);
  });
}

function renderPitchDeck(deck) {
  const container = document.getElementById('slides-container');
  container.innerHTML = '';

  (deck.slides || []).forEach(slide => {
    const card = document.createElement('div');
    card.className = 'slide-box p-4 space-y-3';

    let bulletsHtml = (slide.bullet_points || []).map(b => `
      <li class="flex items-start space-x-2.5">
        <div class="bullet-marker"></div>
        <span class="text-xs text-slate-800 leading-normal">${b}</span>
      </li>
    `).join('');

    card.innerHTML = `
      <div class="flex items-center justify-between border-b border-slate-100 pb-2">
        <div class="flex items-center space-x-2">
          <span class="font-mono text-[10px] uppercase font-bold px-1.5 py-0.5 bg-slate-100 text-slate-700 border border-slate-200">
            Slide 0${slide.slide_number}
          </span>
          <h4 class="text-xs font-bold text-slate-900">${slide.title}</h4>
        </div>
      </div>
      
      <ul class="space-y-1.5 pt-0.5">
        ${bulletsHtml}
      </ul>

      <!-- Speaker Notes Accordion -->
      <div class="pt-2 border-t border-slate-100">
        <button 
          onclick="toggleSpeakerNotes(${slide.slide_number})" 
          class="flex items-center justify-between w-full text-[11px] font-semibold text-slate-600 hover:text-slate-900 py-0.5 transition"
        >
          <span>Presenter Script (Speaker Notes)</span>
          <span id="speaker-icon-${slide.slide_number}" class="font-mono text-[10px]">+</span>
        </button>
        <div id="speaker-notes-${slide.slide_number}" class="speaker-notes-content bg-slate-50 rounded-sm p-2.5 text-xs text-slate-700 mt-1 border border-slate-200 font-mono leading-relaxed">
          ${slide.speaker_notes}
        </div>
      </div>
    `;
    container.appendChild(card);
  });
}

function toggleSpeakerNotes(slideNum) {
  const content = document.getElementById(`speaker-notes-${slideNum}`);
  const icon = document.getElementById(`speaker-icon-${slideNum}`);
  if (content.classList.contains('expanded')) {
    content.classList.remove('expanded');
    icon.textContent = '+';
  } else {
    content.classList.add('expanded');
    icon.textContent = '-';
  }
}

function renderComplianceAudit(audit) {
  const tbody = document.getElementById('claims-table-body');
  tbody.innerHTML = '';

  const confBadge = document.getElementById('audit-deck-conf-badge');
  const confPct = Math.round(audit.deck_confidence_score * 100);
  confBadge.textContent = `Confidence: ${confPct}%`;

  (audit.claims || []).forEach(c => {
    const isVerified = c.status.toLowerCase() === 'verified';
    const statusClass = isVerified 
      ? 'bg-emerald-50 text-emerald-800 border-emerald-200' 
      : 'bg-rose-50 text-rose-800 border-rose-200';
    const statusText = isVerified ? 'VERIFIED' : 'FLAGGED';

    const tr = document.createElement('tr');
    tr.className = 'border-b border-slate-200 text-xs';
    tr.innerHTML = `
      <td class="py-2.5 px-3 whitespace-nowrap align-top">
        <span class="px-2 py-0.5 rounded-full text-[10px] font-bold border ${statusClass}">
          ${statusText}
        </span>
      </td>
      <td class="py-2.5 px-3 font-semibold text-slate-900 align-top">
        ${c.core_policy_feature}
        <div class="text-[10px] font-normal text-slate-500 mt-0.5">Limit: ${c.stated_limit_or_rule}</div>
      </td>
      <td class="py-2.5 px-3 text-slate-700 align-top leading-normal">
        "${c.claim}"
      </td>
      <td class="py-2.5 px-3 text-slate-600 align-top leading-normal">
        ${c.evidence_snippet}
      </td>
      <td class="py-2.5 px-3 font-mono text-[11px] text-slate-500 align-top whitespace-nowrap">
        ${c.source_document.replace('.pdf', '')}
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function renderRAGEvidence(chunks) {
  const container = document.getElementById('chunks-container');
  container.innerHTML = '';

  if (!chunks || chunks.length === 0) {
    container.innerHTML = '<div class="text-xs text-slate-400 font-mono">No policy excerpts retrieved.</div>';
    return;
  }

  chunks.forEach((chunk, idx) => {
    const tier = chunk.relevance_tier || 'High';
    const score = chunk.normalized_score || 85.0;

    let tierBadgeClass = 'bg-emerald-50 text-emerald-800 border-emerald-200';
    if (tier === 'Medium') tierBadgeClass = 'bg-blue-50 text-blue-800 border-blue-200';
    if (tier === 'Low') tierBadgeClass = 'bg-amber-50 text-amber-800 border-amber-200';

    const div = document.createElement('div');
    div.className = 'bg-white border-l-2 border-l-[#00205B] border-y border-r border-slate-200 p-3 space-y-2';
    div.innerHTML = `
      <div class="flex items-center justify-between text-xs">
        <div class="font-medium text-slate-800">
          <span>Passage ${idx + 1}:</span>
          <span class="font-mono text-[11px] text-slate-600 font-semibold ml-1">${chunk.source_document}</span>
          <span class="text-slate-400 ml-1">(Page ${chunk.page})</span>
        </div>
        <span class="px-2 py-0.5 rounded text-[10px] font-bold border ${tierBadgeClass}">
          ${tier} (${score}%)
        </span>
      </div>

      <pre class="bg-slate-50 p-2.5 rounded-sm text-xs font-mono text-slate-700 whitespace-pre-wrap leading-relaxed max-h-40 overflow-y-auto border border-slate-200">${chunk.content}</pre>
    `;
    container.appendChild(div);
  });
}

function renderDownloadsHub(downloads) {
  if (!downloads) return;

  const pdfBtn = document.getElementById('download-pdf-btn');
  const pptxBtn = document.getElementById('download-pptx-btn');
  const csvBtn = document.getElementById('download-csv-btn');
  const jsonBtn = document.getElementById('download-json-btn');

  if (downloads.pdf) {
    pdfBtn.href = downloads.pdf.url;
    pdfBtn.setAttribute('download', downloads.pdf.filename);
  }
  if (downloads.pptx) {
    pptxBtn.href = downloads.pptx.url;
    pptxBtn.setAttribute('download', downloads.pptx.filename);
  }
  if (downloads.csv) {
    csvBtn.href = downloads.csv.url;
    csvBtn.setAttribute('download', downloads.csv.filename);
  }
  if (downloads.json) {
    jsonBtn.href = downloads.json.url;
    jsonBtn.setAttribute('download', downloads.json.filename);
  }
}

// =============================================================================
// 4. Tab Navigation Logic
// =============================================================================
function switchTab(tabId) {
  activeTab = tabId;

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.remove('active', 'text-slate-900');
    btn.classList.add('text-slate-500');
  });

  const activeBtn = document.getElementById(`tab-btn-${tabId}`);
  if (activeBtn) {
    activeBtn.classList.add('active', 'text-slate-900');
    activeBtn.classList.remove('text-slate-500');
  }

  document.querySelectorAll('.tab-panel').forEach(panel => {
    panel.classList.add('hidden');
  });

  const activeContent = document.getElementById(`tab-content-${tabId}`);
  if (activeContent) {
    activeContent.classList.remove('hidden');
  }
}

// =============================================================================
// 5. Notifications
// =============================================================================
function showError(msg) {
  errorMessageEl.textContent = msg;
  errorBanner.classList.remove('hidden');
}

function dismissError() {
  errorBanner.classList.add('hidden');
}
