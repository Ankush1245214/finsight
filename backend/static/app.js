const fileInput = document.querySelector('#file-input');
const dropZone = document.querySelector('#drop-zone');
const browseButton = document.querySelector('#browse-button');
const selectedFilePanel = document.querySelector('#selected-file');
const selectedName = document.querySelector('#selected-name');
const selectedSize = document.querySelector('#selected-size');
const removeFileButton = document.querySelector('#remove-file');
const extractButton = document.querySelector('#extract-button');
const statusMessage = document.querySelector('#status-message');
const uploadHint = document.querySelector('#upload-hint');
const resultsCard = document.querySelector('#results-card');
const resultFilename = document.querySelector('#result-filename');
const pageCount = document.querySelector('#page-count');
const pageNav = document.querySelector('#page-nav');
const activePageLabel = document.querySelector('#active-page-label');
const pageText = document.querySelector('#page-text');
const copyAllButton = document.querySelector('#copy-all');
const copyPageButton = document.querySelector('#copy-page');
const textWarning = document.querySelector('#text-warning');
const resultLayout = document.querySelector('#result-layout');
const documentCount = document.querySelector('#document-count');
const savedCount = document.querySelector('#saved-count');
const savedList = document.querySelector('#saved-list');
const toast = document.querySelector('#toast');
const ocrToggle = document.querySelector('#ocr-toggle');
const documentSearch = document.querySelector('#document-search');
const searchResults = document.querySelector('#search-results');
const exportFormat = document.querySelector('#export-format');
const exportButton = document.querySelector('#export-button');
const ratioList = document.querySelector('#ratio-list');
const ollamaModel = document.querySelector('#ollama-model');
const questionInput = document.querySelector('#question-input');
const askButton = document.querySelector('#ask-button');
const answerBox = document.querySelector('#answer-box');
const answerCitations = document.querySelector('#answer-citations');
const summaryButton = document.querySelector('#summary-button');
const summaryBox = document.querySelector('#summary-box');
const summaryCitations = document.querySelector('#summary-citations');
const annotationPage = document.querySelector('#annotation-page');
const noteInput = document.querySelector('#note-input');
const saveNoteButton = document.querySelector('#save-note');
const highlightSelectionButton = document.querySelector('#highlight-selection');
const annotationList = document.querySelector('#annotation-list');
const tableList = document.querySelector('#table-list');
const exportTablesButton = document.querySelector('#export-tables');
const demoButton = document.querySelector('#demo-button');
const compareFirst = document.querySelector('#compare-first');
const compareSecond = document.querySelector('#compare-second');
const compareButton = document.querySelector('#compare-button');
const compareResults = document.querySelector('#compare-results');
const searchDocument = document.querySelector('#search-document');
const searchPage = document.querySelector('#search-page');
const searchStartDate = document.querySelector('#search-start-date');
const searchEndDate = document.querySelector('#search-end-date');
const exportSelectedPages = document.querySelector('#export-selected-pages');
const exportExtras = document.querySelector('#export-extras');

let selectedFile = null;
let extractedPages = [];
let activePage = 0;
let activeDocumentId = null;
let activeSearchTerm = '';
let documentAnnotations = [];
let selectedPagesForExport = new Set();
let toastTimer;

function formatBytes(bytes) {
  return bytes < 1024 * 1024
    ? `${Math.max(1, Math.round(bytes / 1024))} KB`
    : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function showStatus(message, type = 'info') {
  statusMessage.textContent = message;
  statusMessage.className = `status-message ${type}`;
  statusMessage.hidden = false;
}

function clearStatus() {
  statusMessage.textContent = '';
  statusMessage.hidden = true;
}

function chooseFile(file) {
  clearStatus();
  resultsCard.hidden = true;
  if (!file) return;
  const extension = file.name.split('.').pop()?.toLowerCase();
  if (!['pdf', 'docx', 'xlsx', 'xlsm', 'csv'].includes(extension)) {
    selectedFile = null;
    fileInput.value = '';
    selectedFilePanel.hidden = true;
    extractButton.disabled = true;
    showStatus('Choose a PDF, DOCX, XLSX, XLSM, or CSV file.', 'error');
    return;
  }
  if (file.size > 10 * 1024 * 1024) {
    selectedFile = null;
    fileInput.value = '';
    selectedFilePanel.hidden = true;
    extractButton.disabled = true;
    showStatus('This file is larger than 10 MB. Choose a smaller file to continue.', 'error');
    return;
  }
  selectedFile = file;
  ocrToggle.disabled = extension !== 'pdf';
  if (extension !== 'pdf') ocrToggle.checked = false;
  selectedName.textContent = file.name;
  selectedSize.textContent = formatBytes(file.size);
  selectedFilePanel.hidden = false;
  dropZone.classList.add('has-file');
  extractButton.disabled = false;
  uploadHint.textContent = 'Ready when you are. Text and tables will be extracted locally.';
}

function clearFile() {
  selectedFile = null;
  ocrToggle.disabled = false;
  fileInput.value = '';
  selectedFilePanel.hidden = true;
  dropZone.classList.remove('has-file');
  extractButton.disabled = true;
  uploadHint.textContent = 'Your file is used to extract text for this review.';
  clearStatus();
}

function setLoading(isLoading) {
  extractButton.disabled = isLoading || !selectedFile;
  fileInput.disabled = isLoading;
  browseButton.disabled = isLoading;
  removeFileButton.disabled = isLoading;
  dropZone.setAttribute('aria-disabled', String(isLoading));
  dropZone.classList.toggle('is-loading', isLoading);
  extractButton.classList.toggle('is-loading', isLoading);
  extractButton.querySelector('span').textContent = isLoading ? 'Reading report…' : 'Extract text';
  extractButton.querySelector('svg').innerHTML = isLoading
    ? '<path d="M16.5 10a6.5 6.5 0 1 1-2-4.7"/>'
    : '<path d="M4 10h11m-4-4 4 4-4 4"/>';
  if (isLoading) showStatus('Reading your report. Larger files may take a few seconds.', 'info');
  else if (statusMessage.textContent.startsWith('Reading your report')) clearStatus();
}

async function extractSelectedFile() {
  if (!selectedFile) return;
  setLoading(true);
  const formData = new FormData();
  formData.append('file', selectedFile);
  try {
    const ocrParam = ocrToggle.checked ? '?ocr=true' : '';
    const response = await fetch(`/documents/extract${ocrParam}`, { method: 'POST', body: formData });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) {
      const serverMessage = typeof result.detail === 'string' ? result.detail : 'The report could not be processed.';
      throw new Error(serverMessage);
    }
    renderResults(result);
    showStatus('Text extraction is complete. Select a page to review it.', 'success');
    await loadSavedReports();
    resultsCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    const message = error instanceof TypeError
      ? 'FinSight could not reach the server. Make sure the backend is running, then try again.'
      : error.message;
    showStatus(message, 'error');
  } finally {
    setLoading(false);
  }
}

function renderResults(result) {
  activeDocumentId = result.document_id ?? null;
  extractedPages = Array.isArray(result.pages) ? result.pages : [];
  activePage = 0;
  activeSearchTerm = result.search_term || '';
  selectedPagesForExport.clear();
  answerBox.hidden = true;
  answerCitations.hidden = true;
  summaryBox.hidden = true;
  summaryCitations.hidden = true;
  resultsCard.hidden = false;
  resultFilename.textContent = result.filename || selectedFile?.name || 'Uploaded PDF';
  pageCount.textContent = `${result.page_count ?? extractedPages.length} ${result.page_count === 1 ? 'page' : 'pages'}`;
  const hasText = extractedPages.some((page) => page.text && page.text.trim());
  textWarning.hidden = hasText;
  resultLayout.hidden = !hasText;
  pageNav.replaceChildren();
  extractedPages.forEach((page) => selectedPagesForExport.add(page.page_number));
  if (!hasText) {
    loadRatios();
    loadAnnotations();
    loadTables();
    return;
  }
  extractedPages.forEach((page, index) => {
    const item = document.createElement('div');
    item.className = 'page-nav-item';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = true;
    checkbox.setAttribute('aria-label', `Select page ${page.page_number} for export`);
    checkbox.addEventListener('change', () => {
      if (checkbox.checked) selectedPagesForExport.add(page.page_number);
      else selectedPagesForExport.delete(page.page_number);
    });
    const button = document.createElement('button');
    button.className = 'page-tab';
    button.type = 'button';
    button.setAttribute('aria-current', index === activePage ? 'page' : 'false');
    const title = document.createElement('span');
    title.textContent = `Page ${page.page_number}`;
    const size = document.createElement('small');
    size.textContent = page.text.trim() ? `${page.text.trim().length} chars` : 'No text';
    button.append(title, size);
    button.addEventListener('click', () => selectPage(index));
    item.append(checkbox, button);
    pageNav.append(item);
  });
  updateActivePage();
  loadRatios();
  loadAnnotations();
  loadTables();
}

function selectPage(index) {
  activePage = index;
  updateActivePage();
}

function updateActivePage() {
  const page = extractedPages[activePage];
  if (!page) return;
  activePageLabel.textContent = `PAGE ${String(page.page_number).padStart(2, '0')}`;
  annotationPage.textContent = `PAGE ${page.page_number}`;
  renderPageText(page.text || 'No text was found on this page.');
  renderAnnotations();
  [...pageNav.querySelectorAll('.page-tab')].forEach((button, index) => {
    button.setAttribute('aria-current', index === activePage ? 'page' : 'false');
  });
}

function renderPageText(text) {
  const terms = [
    activeSearchTerm,
    ...documentAnnotations
      .filter((item) => item.page_number === extractedPages[activePage]?.page_number && item.kind === 'highlight')
      .map((item) => item.selected_text),
  ].filter(Boolean);
  const ranges = [];
  terms.forEach((term, termIndex) => {
    let start = 0;
    while (start < text.length) {
      const index = text.toLowerCase().indexOf(term.toLowerCase(), start);
      if (index < 0) break;
      ranges.push({ start: index, end: index + term.length, search: termIndex === 0 && activeSearchTerm });
      start = index + Math.max(term.length, 1);
    }
  });
  ranges.sort((a, b) => a.start - b.start || b.end - a.end);
  pageText.replaceChildren();
  let cursor = 0;
  for (const range of ranges) {
    if (range.start < cursor) continue;
    pageText.append(document.createTextNode(text.slice(cursor, range.start)));
    const mark = document.createElement('mark');
    mark.className = range.search ? 'search-highlight' : 'saved-highlight';
    mark.textContent = text.slice(range.start, range.end);
    pageText.append(mark);
    cursor = range.end;
  }
  pageText.append(document.createTextNode(text.slice(cursor)));
}

async function loadAnnotations() {
  if (!activeDocumentId) {
    documentAnnotations = [];
    annotationList.replaceChildren();
    return;
  }
  try {
    const response = await fetch(`/documents/${activeDocumentId}/annotations`);
    if (!response.ok) throw new Error('Notes could not be loaded.');
    const result = await response.json();
    documentAnnotations = result.annotations || [];
    updateActivePage();
  } catch (error) {
    showToast(error.message);
  }
}

function renderAnnotations() {
  if (!annotationList) return;
  const pageNumber = extractedPages[activePage]?.page_number;
  const items = documentAnnotations.filter((item) => item.page_number === pageNumber);
  annotationList.replaceChildren();
  if (!items.length) {
    const empty = document.createElement('span');
    empty.className = 'insight-empty';
    empty.textContent = 'No notes or highlights on this page yet.';
    annotationList.append(empty);
    return;
  }
  items.forEach((item) => {
    const row = document.createElement('div');
    row.className = `annotation-item annotation-${item.kind}`;
    const content = document.createElement('span');
    content.textContent = item.kind === 'highlight'
      ? `Highlighted: ${item.selected_text}${item.body ? ` — ${item.body}` : ''}`
      : item.body;
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.setAttribute('aria-label', 'Delete annotation');
    remove.textContent = '×';
    remove.addEventListener('click', () => deleteAnnotation(item.id));
    row.append(content, remove);
    annotationList.append(row);
  });
}

async function copyText(value, message) {
  if (!value) {
    showToast('There is no text to copy on this page.');
    return;
  }
  try {
    await navigator.clipboard.writeText(value);
    showToast(message);
  } catch {
    const temporary = document.createElement('textarea');
    temporary.value = value;
    temporary.style.position = 'fixed';
    temporary.style.opacity = '0';
    document.body.append(temporary);
    temporary.select();
    const copied = document.execCommand('copy');
    temporary.remove();
    showToast(copied ? message : 'Could not copy text. Select it and copy manually.');
  }
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add('is-visible');
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toast.classList.remove('is-visible'), 2300);
}

async function loadRatios() {
  if (!activeDocumentId) {
    ratioList.innerHTML = '<span class="insight-empty">Open a saved report to calculate ratios.</span>';
    return;
  }
  try {
    const response = await fetch(`/documents/${activeDocumentId}/ratios`);
    if (!response.ok) throw new Error('Ratios could not be calculated.');
    const result = await response.json();
    const labels = {
      debt_to_assets: 'Liabilities / assets',
      net_profit_margin: 'Net margin',
      equity_to_assets: 'Equity / assets',
      gross_margin: 'Gross margin',
      operating_margin: 'Operating margin',
      current_ratio: 'Current ratio',
      operating_cash_flow_ratio: 'Operating cash flow / current liabilities',
      free_cash_flow_margin: 'Free cash flow margin',
    };
    ratioList.replaceChildren();
    const entries = Object.entries(result.ratios);
    if (!entries.length) {
      const empty = document.createElement('span');
      empty.className = 'insight-empty';
      empty.textContent = 'No recognizable financial values were found.';
      ratioList.append(empty);
      return;
    }
    entries.forEach(([key, value]) => {
      const item = document.createElement('div');
      item.className = 'ratio-item';
      const strong = document.createElement('strong');
      strong.textContent = key === 'current_ratio' ? `${value.toFixed(2)}×` : `${(value * 100).toFixed(1)}%`;
      const label = document.createElement('span');
      label.textContent = labels[key] || key;
      const explanation = document.createElement('small');
      explanation.textContent = result.explanations?.[key] || '';
      item.title = explanation.textContent;
      item.append(strong, label, explanation);
      ratioList.append(item);
    });
  } catch (error) {
    ratioList.innerHTML = '<span class="insight-empty">Ratios are unavailable.</span>';
  }
}

async function loadTables() {
  if (!activeDocumentId) return;
  try {
    const response = await fetch(`/documents/${activeDocumentId}/tables`);
    if (!response.ok) throw new Error('Extracted tables could not be loaded.');
    const result = await response.json();
    tableList.replaceChildren();
    if (!result.tables.length) {
      const empty = document.createElement('span');
      empty.className = 'insight-empty';
      empty.textContent = 'No tables were detected in this report.';
      tableList.append(empty);
      return;
    }
    result.tables.forEach((table) => {
      const section = document.createElement('section');
      section.className = 'extracted-table';
      const title = document.createElement('strong');
      title.textContent = `${table.title} · page ${table.page_number}`;
      const wrapper = document.createElement('div');
      wrapper.className = 'table-scroll';
      const htmlTable = document.createElement('table');
      table.rows.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        row.forEach((value) => {
          const cell = document.createElement(rowIndex === 0 ? 'th' : 'td');
          cell.textContent = value ?? '';
          tr.append(cell);
        });
        htmlTable.append(tr);
      });
      wrapper.append(htmlTable);
      section.append(title, wrapper);
      tableList.append(section);
    });
  } catch (error) {
    tableList.textContent = error.message;
  }
}

let searchTimer;
function scheduleSearch() {
  window.clearTimeout(searchTimer);
  searchTimer = window.setTimeout(searchDocuments, 180);
}
documentSearch?.addEventListener('input', scheduleSearch);
[searchDocument, searchPage, searchStartDate, searchEndDate].forEach((input) => {
  input?.addEventListener('change', scheduleSearch);
});

async function searchDocuments() {
  const query = documentSearch.value.trim();
  if (!query) {
    searchResults.hidden = true;
    searchResults.replaceChildren();
    return;
  }
  try {
    const params = new URLSearchParams({ q: query });
    if (searchDocument.value) params.set('document_id', searchDocument.value);
    if (searchPage.value) params.set('page_number', searchPage.value);
    if (searchStartDate.value) params.set('start_date', searchStartDate.value);
    if (searchEndDate.value) params.set('end_date', searchEndDate.value);
    const response = await fetch(`/documents/search?${params}`);
    if (!response.ok) throw new Error();
    const result = await response.json();
    searchResults.replaceChildren();
    if (!result.matches.length && !result.documents.length) {
      const empty = document.createElement('div');
      empty.className = 'search-result';
      empty.textContent = 'No saved reports match this search.';
      searchResults.append(empty);
    } else {
      const hits = result.matches.length ? result.matches : result.documents;
      hits.forEach((saved) => {
        const row = document.createElement('div');
        row.className = 'search-result';
        const open = document.createElement('button');
        open.type = 'button';
        const documentId = saved.document_id ?? saved.id;
        const label = saved.page_number
          ? `${saved.filename} · page ${saved.page_number}`
          : saved.filename;
        open.textContent = label;
        open.addEventListener('click', () => openSavedReport(documentId, saved.page_number, query));
        row.append(open);
        if (saved.text) {
          const excerpt = document.createElement('p');
          const index = saved.text.toLowerCase().indexOf(query.toLowerCase());
          const start = Math.max(0, index - 70);
          excerpt.textContent = `${start ? '…' : ''}${saved.text.slice(start, start + 180).trim()}…`;
          row.append(excerpt);
        }
        searchResults.append(row);
      });
    }
    searchResults.hidden = false;
  } catch {
    searchResults.replaceChildren();
    const error = document.createElement('div');
    error.className = 'search-result';
    error.textContent = 'Search is unavailable. Check the local server and try again.';
    searchResults.append(error);
    searchResults.hidden = false;
  }
}

exportButton?.addEventListener('click', () => {
  if (!activeDocumentId) {
    showToast('Open a saved report before exporting.');
    return;
  }
  const params = new URLSearchParams({ format: exportFormat.value });
  if (exportSelectedPages.checked) {
    if (!selectedPagesForExport.size) {
      showToast('Select at least one page to export.');
      return;
    }
    params.set('pages', [...selectedPagesForExport].sort((a, b) => a - b).join(','));
  }
  if (exportExtras.checked) params.set('include_annotations', 'true');
  window.location.href = `/documents/${activeDocumentId}/export?${params}`;
});

askButton?.addEventListener('click', async () => {
  if (!activeDocumentId) {
    showToast('Open a saved report before asking a question.');
    return;
  }
  const question = questionInput.value.trim();
  if (!question) {
    showToast('Enter a question first.');
    return;
  }
  askButton.disabled = true;
  askButton.textContent = 'Thinking…';
  answerBox.hidden = false;
  answerBox.textContent = 'Connecting to local Ollama…';
  try {
    const response = await fetch(`/documents/${activeDocumentId}/ask`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question, model: ollamaModel.value.trim() || 'llama3.2' }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || 'Ollama could not answer.');
    answerBox.textContent = result.answer;
    renderCitations(answerCitations, result.citations);
  } catch (error) {
    answerBox.textContent = error.message;
    answerCitations.hidden = true;
  } finally {
    askButton.disabled = false;
    askButton.textContent = 'Ask';
  }
});

questionInput?.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') askButton.click();
});

summaryButton?.addEventListener('click', async () => {
  if (!activeDocumentId) {
    showToast('Open a saved report before generating a summary.');
    return;
  }
  summaryButton.disabled = true;
  summaryButton.textContent = 'Summarizing…';
  summaryBox.hidden = false;
  summaryBox.textContent = 'Generating a cited review with local Ollama…';
  try {
    const response = await fetch(`/documents/${activeDocumentId}/summary`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question: 'summary', model: ollamaModel.value.trim() || 'llama3.2' }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || 'The report summary could not be generated.');
    summaryBox.textContent = result.summary;
    renderCitations(summaryCitations, result.citations);
  } catch (error) {
    summaryBox.textContent = error.message;
    summaryCitations.hidden = true;
  } finally {
    summaryButton.disabled = false;
    summaryButton.textContent = 'Generate report summary';
  }
});

function renderCitations(container, citations = []) {
  container.replaceChildren();
  if (!citations.length) {
    container.hidden = true;
    return;
  }
  const label = document.createElement('span');
  label.textContent = 'Sources: ';
  container.append(label);
  citations.forEach((citation) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = `Page ${citation.page_number}`;
    button.title = citation.excerpt || `Open page ${citation.page_number}`;
    button.addEventListener('click', () => goToPage(citation.page_number));
    container.append(button);
  });
  container.hidden = false;
}

function goToPage(pageNumber) {
  const pageIndex = extractedPages.findIndex((page) => page.page_number === pageNumber);
  if (pageIndex >= 0) {
    selectPage(pageIndex);
    resultsCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

async function saveAnnotation(kind, selectedText = '', body = '') {
  if (!activeDocumentId) {
    showToast('Open a saved report before adding notes or highlights.');
    return;
  }
  const pageNumber = extractedPages[activePage]?.page_number;
  try {
    const response = await fetch(`/documents/${activeDocumentId}/annotations`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        page_number: pageNumber,
        kind,
        selected_text: selectedText,
        body,
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || 'The annotation could not be saved.');
    documentAnnotations.push(result.annotation);
    updateActivePage();
    if (kind === 'note') noteInput.value = '';
    showToast(kind === 'note' ? 'Note saved to this page.' : 'Highlight saved to this page.');
  } catch (error) {
    showToast(error.message);
  }
}

saveNoteButton?.addEventListener('click', () => saveAnnotation('note', '', noteInput.value));
noteInput?.addEventListener('keydown', (event) => {
  if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') saveNoteButton.click();
});
highlightSelectionButton?.addEventListener('click', () => {
  const selection = window.getSelection();
  const selectedText = selection?.toString().trim() || '';
  if (!selectedText || !pageText.contains(selection.anchorNode)) {
    showToast('Select text in the page first.');
    return;
  }
  saveAnnotation('highlight', selectedText);
  selection.removeAllRanges();
});

async function deleteAnnotation(annotationId) {
  try {
    const response = await fetch(`/documents/${activeDocumentId}/annotations/${annotationId}`, {
      method: 'DELETE',
    });
    if (!response.ok) throw new Error('The annotation could not be deleted.');
    documentAnnotations = documentAnnotations.filter((item) => item.id !== annotationId);
    updateActivePage();
  } catch (error) {
    showToast(error.message);
  }
}

exportTablesButton?.addEventListener('click', () => {
  if (!activeDocumentId) {
    showToast('Open a saved report before exporting tables.');
    return;
  }
  window.location.href = `/documents/${activeDocumentId}/tables/export`;
});

function formatSavedDate(value) {
  const date = new Date(`${value}Z`);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
}

async function loadSavedReports() {
  try {
    const response = await fetch('/documents');
    if (!response.ok) throw new Error('Saved reports could not be loaded.');
    const result = await response.json();
    renderSavedReports(result.documents || []);
    updateReportSelectors(result.documents || []);
  } catch (error) {
    documentCount.textContent = '0';
    savedCount.textContent = 'Unavailable';
  }
}

function renderSavedReports(documents) {
  documentCount.textContent = String(documents.length);
  savedCount.textContent = `${documents.length} ${documents.length === 1 ? 'report' : 'reports'}`;
  savedList.replaceChildren();
  if (!documents.length) {
    const empty = document.createElement('p');
    empty.className = 'saved-empty';
    empty.textContent = 'Reports you extract will appear here for later review.';
    savedList.append(empty);
    return;
  }
  documents.forEach((savedDocument) => {
    const row = document.createElement('div');
    row.className = 'saved-report';
    const icon = document.createElement('span');
    icon.className = 'saved-report-icon';
    icon.textContent = (savedDocument.source_format || 'pdf').toUpperCase().slice(0, 4);
    const info = document.createElement('div');
    info.className = 'saved-report-info';
    const open = document.createElement('button');
    open.type = 'button';
    open.textContent = savedDocument.filename;
    open.title = `Open ${savedDocument.filename}`;
    open.addEventListener('click', () => openSavedReport(savedDocument.id));
    const meta = document.createElement('small');
    meta.textContent = `${savedDocument.page_count} ${savedDocument.page_count === 1 ? 'page' : 'pages'} · ${(savedDocument.source_format || 'pdf').toUpperCase()} · ${formatSavedDate(savedDocument.created_at)}`;
    info.append(open, meta);
    const remove = document.createElement('button');
    remove.className = 'delete-report';
    remove.type = 'button';
    remove.textContent = '×';
    remove.title = `Delete ${savedDocument.filename}`;
    remove.setAttribute('aria-label', `Delete ${savedDocument.filename}`);
    remove.addEventListener('click', () => deleteSavedReport(savedDocument.id, savedDocument.filename));
    row.append(icon, info, remove);
    savedList.append(row);
  });
}

function updateReportSelectors(documents) {
  const selectors = [searchDocument, compareFirst, compareSecond];
  selectors.forEach((select) => {
    const previous = select.value;
    select.replaceChildren();
    const empty = document.createElement('option');
    empty.value = '';
    empty.textContent = select === searchDocument ? 'All reports' : 'Choose a report';
    select.append(empty);
    documents.forEach((savedDocument) => {
      const option = document.createElement('option');
      option.value = savedDocument.id;
      option.textContent = savedDocument.filename;
      select.append(option);
    });
    if ([...select.options].some((option) => option.value === previous)) select.value = previous;
  });
  if (!compareFirst.value && documents.length) compareFirst.value = String(documents.at(-1).id);
  if (!compareSecond.value && documents.length > 1) compareSecond.value = String(documents[0].id);
}

async function openSavedReport(documentId, pageNumber = null, searchTerm = '') {
  try {
    const response = await fetch(`/documents/${documentId}`);
    if (!response.ok) throw new Error('The saved report could not be opened.');
    const document = await response.json();
    renderResults({
      document_id: documentId,
      filename: document.filename,
      page_count: document.page_count,
      pages: document.pages,
      search_term: searchTerm,
    });
    if (pageNumber) {
      const index = extractedPages.findIndex((page) => page.page_number === pageNumber);
      if (index >= 0) selectPage(index);
    }
    showStatus('Saved report opened from local history.', 'success');
    resultsCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    showToast(error.message);
  }
}

compareButton?.addEventListener('click', async () => {
  if (!compareFirst.value || !compareSecond.value) {
    showToast('Choose two saved reports to compare.');
    return;
  }
  compareButton.disabled = true;
  compareResults.hidden = false;
  compareResults.textContent = 'Comparing financial figures…';
  try {
    const params = new URLSearchParams({
      first_id: compareFirst.value,
      second_id: compareSecond.value,
    });
    const response = await fetch(`/documents/compare?${params}`);
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || 'Reports could not be compared.');
    renderComparison(result);
  } catch (error) {
    compareResults.textContent = error.message;
  } finally {
    compareButton.disabled = false;
  }
});

function renderComparison(result) {
  compareResults.replaceChildren();
  const heading = document.createElement('p');
  heading.className = 'compare-caption';
  heading.textContent = `${result.first.filename} → ${result.second.filename}`;

  const charts = document.createElement('div');
  charts.className = 'comparison-charts';
  const financialMetrics = [
    ['revenue', 'Revenue'],
    ['gross_profit', 'Gross profit'],
    ['operating_income', 'Operating income'],
    ['net_income', 'Net income'],
    ['assets', 'Total assets'],
    ['liabilities', 'Total liabilities'],
    ['equity', 'Total equity'],
  ].map(([key, label]) => ({ key, label, ...result.metrics[key] }))
    .filter((metric) => metric.first !== null && metric.first !== undefined
      && metric.second !== null && metric.second !== undefined);
  const ratios = [
    ['gross_margin', 'Gross margin', '%'],
    ['operating_margin', 'Operating margin', '%'],
    ['net_profit_margin', 'Net profit margin', '%'],
    ['current_ratio', 'Current ratio', '×'],
    ['debt_to_assets', 'Liabilities / assets', '%'],
    ['equity_to_assets', 'Equity / assets', '%'],
  ].map(([key, label, unit]) => ({
    key,
    label,
    unit,
    first: result.first.ratios[key],
    second: result.second.ratios[key],
  })).filter((metric) => metric.first !== undefined && metric.second !== undefined);

  if (financialMetrics.length) {
    const section = document.createElement('section');
    section.className = 'chart-card';
    const title = document.createElement('h4');
    title.textContent = 'Financial figures';
    const description = document.createElement('p');
    description.className = 'chart-description';
    description.textContent = 'Grouped bars compare extracted amounts across the selected reports.';
    section.append(title, description, createFinancialBarChart(
      financialMetrics,
      result.first.filename,
      result.second.filename,
    ));
    charts.append(section);
  }
  if (ratios.length) {
    const section = document.createElement('section');
    section.className = 'chart-card';
    const title = document.createElement('h4');
    title.textContent = 'Ratio comparison';
    const description = document.createElement('p');
    description.className = 'chart-description';
    description.textContent = 'Each row compares the same ratio; values show percentages or multiples as appropriate.';
    section.append(title, description, createRatioChart(
      ratios,
      result.first.filename,
      result.second.filename,
    ));
    charts.append(section);
  }

  const table = document.createElement('table');
  table.className = 'comparison-data-table';
  const head = document.createElement('tr');
  ['Metric', 'First', 'Second', 'Change'].forEach((label) => {
    const cell = document.createElement('th');
    cell.textContent = label;
    head.append(cell);
  });
  table.append(head);
  const labels = {
    assets: 'Total assets',
    liabilities: 'Total liabilities',
    equity: 'Total equity',
    revenue: 'Revenue',
    net_income: 'Net income',
    gross_profit: 'Gross profit',
    operating_income: 'Operating income',
    current_assets: 'Current assets',
    current_liabilities: 'Current liabilities',
    cash_from_operations: 'Cash from operations',
    capital_expenditures: 'Capital expenditures',
    free_cash_flow: 'Free cash flow',
  };
  const ratioLabels = {
    debt_to_assets: 'Liabilities / assets',
    net_profit_margin: 'Net profit margin',
    equity_to_assets: 'Equity / assets',
    gross_margin: 'Gross margin',
    operating_margin: 'Operating margin',
    current_ratio: 'Current ratio',
    operating_cash_flow_ratio: 'Operating cash flow ratio',
    free_cash_flow_margin: 'Free cash flow margin',
  };
  Object.entries(result.metrics).forEach(([key, metric]) => {
    const row = document.createElement('tr');
    const values = [
      labels[key] || key.replaceAll('_', ' '),
      formatMetric(metric.first),
      formatMetric(metric.second),
      metric.change_percent === null
        ? '—'
        : `${metric.change > 0 ? '+' : ''}${metric.change_percent.toFixed(1)}%`,
    ];
    values.forEach((value) => {
      const cell = document.createElement('td');
      cell.textContent = value;
      row.append(cell);
    });
    table.append(row);
  });
  const ratioKeys = new Set([
    ...Object.keys(result.first.ratios || {}),
    ...Object.keys(result.second.ratios || {}),
  ]);
  if (ratioKeys.size) {
    const labelRow = document.createElement('tr');
    const label = document.createElement('th');
    label.colSpan = 4;
    label.textContent = 'Ratios';
    labelRow.append(label);
    table.append(labelRow);
    ratioKeys.forEach((key) => {
      const row = document.createElement('tr');
      const name = document.createElement('td');
      name.textContent = ratioLabels[key] || key.replaceAll('_', ' ');
      const first = document.createElement('td');
      first.textContent = formatRatio(result.first.ratios[key], key);
      const second = document.createElement('td');
      second.textContent = formatRatio(result.second.ratios[key], key);
      const change = document.createElement('td');
      const before = result.first.ratios[key];
      const after = result.second.ratios[key];
      change.textContent = before === undefined || after === undefined
        ? '—'
        : key === 'current_ratio'
          ? `${(after - before).toFixed(2)}×`
          : `${((after - before) * 100).toFixed(1)} pp`;
      row.append(name, first, second, change);
      table.append(row);
    });
  }
  const tableDetails = document.createElement('details');
  tableDetails.className = 'comparison-table-details';
  const tableSummary = document.createElement('summary');
  tableSummary.textContent = 'View exact figures and percentage changes';
  tableDetails.append(tableSummary, table);
  compareResults.append(heading, charts, tableDetails);
}

function createSvgChart(titleText, descriptionText, height = 220) {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', `0 0 760 ${height}`);
  svg.setAttribute('role', 'img');
  svg.setAttribute('aria-label', titleText);
  svg.classList.add('finance-chart');
  const title = document.createElementNS(svg.namespaceURI, 'title');
  title.textContent = titleText;
  const description = document.createElementNS(svg.namespaceURI, 'desc');
  description.textContent = descriptionText;
  svg.append(title, description);
  return svg;
}

function appendSvgText(svg, text, x, y, className = '') {
  const node = document.createElementNS(svg.namespaceURI, 'text');
  node.setAttribute('x', String(x));
  node.setAttribute('y', String(y));
  if (className) node.setAttribute('class', className);
  node.textContent = text;
  svg.append(node);
  return node;
}

function appendSvgElement(svg, name, attributes) {
  const node = document.createElementNS(svg.namespaceURI, name);
  Object.entries(attributes).forEach(([key, value]) => node.setAttribute(key, String(value)));
  svg.append(node);
  return node;
}

function createFinancialBarChart(metrics, firstName, secondName) {
  const height = Math.max(220, metrics.length * 38 + 72);
  const svg = createSvgChart(
    'Financial figures comparison',
    metrics.map((item) => `${item.label}: ${firstName} ${formatMetric(item.first)}, ${secondName} ${formatMetric(item.second)}`).join('. '),
    height,
  );
  const labelWidth = 126;
  const plotWidth = 530;
  const top = 38;
  const rowHeight = 34;
  const amounts = metrics.flatMap((item) => [item.first, item.second]);
  const domainMinimum = Math.min(0, ...amounts);
  const domainMaximum = Math.max(0, ...amounts);
  const padding = (domainMaximum - domainMinimum || 1) * 0.08;
  const axisMinimum = domainMinimum - padding;
  const axisMaximum = domainMaximum + padding;
  const axisRange = axisMaximum - axisMinimum;
  const mapValue = (value) => labelWidth + ((value - axisMinimum) / axisRange) * plotWidth;
  const zeroX = mapValue(0);

  [axisMinimum, 0, axisMaximum].forEach((value) => {
    const x = mapValue(value);
    appendSvgElement(svg, 'line', {
      x1: x, y1: top - 8, x2: x, y2: top + metrics.length * rowHeight - 4,
      class: value === 0 ? 'chart-zero-line' : 'chart-grid-line',
    });
    appendSvgText(svg, compactNumber(value), x, top - 14, 'chart-axis-label')
      .setAttribute('text-anchor', value === axisMinimum ? 'start' : value === axisMaximum ? 'end' : 'middle');
  });

  metrics.forEach((metric, index) => {
    const y = top + index * rowHeight;
    appendSvgText(svg, metric.label, labelWidth - 10, y + 12, 'chart-category-label')
      .setAttribute('text-anchor', 'end');
    [
      { value: metric.first, offset: 0, className: 'chart-bar chart-bar-first' },
      { value: metric.second, offset: 13, className: 'chart-bar chart-bar-second' },
    ].forEach((series) => {
      const valueX = mapValue(series.value);
      appendSvgElement(svg, 'rect', {
        x: Math.min(zeroX, valueX),
        y: y + series.offset,
        width: Math.max(1, Math.abs(valueX - zeroX)),
        height: 9,
        rx: 4.5,
        class: series.className,
      });
      appendSvgText(
        svg,
        compactNumber(series.value),
        valueX >= zeroX ? Math.min(valueX + 7, 720) : Math.max(valueX - 7, labelWidth),
        y + series.offset + 8,
        'chart-value-label',
      ).setAttribute('text-anchor', valueX < zeroX ? 'end' : 'start');
    });
  });
  addChartLegend(svg, firstName, secondName);
  return svg;
}

function createRatioChart(ratios, firstName, secondName) {
  const height = Math.max(205, ratios.length * 36 + 88);
  const svg = createSvgChart(
    'Financial ratio comparison',
    ratios.map((item) => `${item.label}: ${firstName} ${formatRatio(item.first, item.key)}, ${secondName} ${formatRatio(item.second, item.key)}`).join('. '),
    height,
  );
  const labelWidth = 150;
  const plotWidth = 510;
  const top = 58;
  const rowHeight = 36;
  ratios.forEach((ratio, index) => {
    const y = top + index * rowHeight;
    appendSvgText(svg, ratio.label, labelWidth - 12, y + 5, 'chart-category-label')
      .setAttribute('text-anchor', 'end');
    const domainMinimum = Math.min(0, ratio.first, ratio.second);
    const domainMaximum = Math.max(0, ratio.first, ratio.second);
    const padding = (domainMaximum - domainMinimum || 1) * 0.1;
    const rowMinimum = domainMinimum - padding;
    const rowRange = domainMaximum - domainMinimum + padding * 2 || 1;
    const mapValue = (value) => labelWidth + ((value - rowMinimum) / rowRange) * plotWidth;
    appendSvgElement(svg, 'line', {
      x1: labelWidth, y1: y, x2: labelWidth + plotWidth, y2: y,
      class: 'chart-track-line',
    });
    const firstX = mapValue(ratio.first);
    const secondX = mapValue(ratio.second);
    appendSvgElement(svg, 'line', {
      x1: firstX, y1: y - 4, x2: secondX, y2: y + 4,
      class: 'chart-dumbbell-line',
    });
    [
      { value: ratio.first, x: firstX, y: y - 4, className: 'chart-point chart-point-first' },
      { value: ratio.second, x: secondX, y: y + 4, className: 'chart-point chart-point-second' },
    ].forEach((series) => {
      const labelX = series.x > labelWidth + plotWidth * 0.78 ? series.x - 9 : series.x + 9;
      appendSvgElement(svg, 'circle', {
        cx: series.x, cy: series.y, r: 5, class: series.className,
      }).setAttribute('aria-label', `${ratio.label}: ${formatRatio(series.value, ratio.key)}`);
      appendSvgText(
        svg,
        formatRatio(series.value, ratio.key),
        labelX,
        series.y + 3,
        'chart-value-label',
      ).setAttribute('text-anchor', labelX < series.x ? 'end' : 'start');
    });
  });
  addChartLegend(svg, firstName, secondName);
  return svg;
}

function compactNumber(value) {
  return new Intl.NumberFormat(undefined, {
    notation: 'compact',
    maximumFractionDigits: 1,
  }).format(value);
}

function addChartLegend(svg, firstName, secondName) {
  const height = Number(svg.getAttribute('viewBox').split(' ')[3]);
  const y = height - 12;
  appendSvgElement(svg, 'circle', { cx: 132, cy: y - 3, r: 4, class: 'chart-point-first' });
  appendSvgText(svg, firstName, 142, y, 'chart-legend-label');
  appendSvgElement(svg, 'circle', { cx: 410, cy: y - 3, r: 4, class: 'chart-point-second' });
  appendSvgText(svg, secondName, 420, y, 'chart-legend-label');
}

function formatMetric(value) {
  return value === null || value === undefined
    ? '—'
    : new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }).format(value);
}

function formatRatio(value, key) {
  if (value === undefined) return '—';
  return key === 'current_ratio' ? `${value.toFixed(2)}×` : `${(value * 100).toFixed(1)}%`;
}

demoButton?.addEventListener('click', async () => {
  demoButton.disabled = true;
  try {
    const response = await fetch('/documents/demo', { method: 'POST' });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || 'Sample reports could not be added.');
    await loadSavedReports();
    compareFirst.value = String(result.documents[0].document_id);
    compareSecond.value = String(result.documents[1].document_id);
    showToast('Two fictional sample reports are ready to explore.');
  } catch (error) {
    showToast(error.message);
  } finally {
    demoButton.disabled = false;
  }
});

async function deleteSavedReport(documentId, filename) {
  if (!window.confirm(`Delete "${filename}" from local history?`)) return;
  try {
    const response = await fetch(`/documents/${documentId}`, { method: 'DELETE' });
    if (!response.ok) throw new Error('The saved report could not be deleted.');
    await loadSavedReports();
    showToast('Saved report deleted.');
  } catch (error) {
    showToast(error.message);
  }
}

browseButton.addEventListener('click', (event) => {
  event.stopPropagation();
  fileInput.click();
});
dropZone.addEventListener('click', (event) => {
  if (event.target.closest('button')) return;
  fileInput.click();
});
fileInput.addEventListener('change', () => chooseFile(fileInput.files?.[0]));
removeFileButton.addEventListener('click', clearFile);
extractButton.addEventListener('click', extractSelectedFile);
copyAllButton.addEventListener('click', () => {
  const allText = extractedPages.map((page) => `Page ${page.page_number}\n${page.text}`).join('\n\n');
  copyText(allText, 'All report text copied.');
});
copyPageButton.addEventListener('click', () => copyText(extractedPages[activePage]?.text, 'Page text copied.'));

for (const eventName of ['dragenter', 'dragover']) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add('is-dragging');
  });
}
for (const eventName of ['dragleave', 'drop']) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove('is-dragging');
  });
}
dropZone.addEventListener('drop', (event) => chooseFile(event.dataTransfer?.files?.[0]));

loadSavedReports();
