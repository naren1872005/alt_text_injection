// PDF Figure & Formula Accessibility Tag Extractor & Excel ALT Manifest Visual Gallery - Frontend Logic

// Global State
let currentSession = null;
let currentFigures = [];         // PDF figures
let currentFormulas = [];        // PDF formulas
let currentExcelRecords = [];    // Excel records
let currentMatchedFigures = [];  // Matched PDF figures
let currentSourceTab = 'pdf';    // 'pdf' | 'formula' | 'excel' | 'match'
let activeFilter = 'all';        // 'all' | 'has_alt' | 'missing' | 'has_image'
let currentView = 'grid';        // 'grid' | 'table'

// Excel Pagination State
let excelPage = 1;
let excelPageSize = 48;
let filteredExcelRecords = [];

// Formula Pagination State
let formulaPage = 1;
let formulaPageSize = 50;
let filteredFormulas = [];
let selectedFormulaIds = new Set();
let currentModalFormula = null;

// DOM Elements: Navigation & Upload
const dropZone = document.getElementById('dropZone');
const pdfFileInput = document.getElementById('pdfFileInput');
const excelFileInput = document.getElementById('excelFileInput');
const browseBtn = document.getElementById('browseBtn');
const browseExcelBtn = document.getElementById('browseExcelBtn');
const samplePdfBtn = document.getElementById('samplePdfBtn');
const sampleExcelBtn = document.getElementById('sampleExcelBtn');
const excelUploadBtn = document.getElementById('excelUploadBtn');

const uploadSection = document.getElementById('uploadSection');
const loadingSection = document.getElementById('loadingSection');
const resultsSection = document.getElementById('resultsSection');
const uploadAnotherBtn = document.getElementById('uploadAnotherBtn');

// Source Switcher Tabs
const tabPdfBtn = document.getElementById('tabPdfBtn');
const tabFormulaBtn = document.getElementById('tabFormulaBtn');
const tabExcelBtn = document.getElementById('tabExcelBtn');
const tabMatchBtn = document.getElementById('tabMatchBtn');
const tabPdfBadge = document.getElementById('tabPdfBadge');
const tabFormulaBadge = document.getElementById('tabFormulaBadge');
const tabExcelBadge = document.getElementById('tabExcelBadge');
const tabMatchBadge = document.getElementById('tabMatchBadge');

// Header Info
const docBadge = document.getElementById('docBadge');
const docFilename = document.getElementById('docFilename');
const docMeta = document.getElementById('docMeta');
const downloadUnselectedBtn = document.getElementById('downloadUnselectedBtn');
const downloadUnselectedLabel = document.getElementById('downloadUnselectedLabel');
const downloadFormulasZipBtn = document.getElementById('downloadFormulasZipBtn');
const downloadExcelZipBtn = document.getElementById('downloadExcelZipBtn');
const exportJsonBtn = document.getElementById('exportJsonBtn');
const injectAltBtn = document.getElementById('injectAltBtn');
const injectAltBtnText = document.getElementById('injectAltBtnText');
const downloadInjectedPdfBtn = document.getElementById('downloadInjectedPdfBtn');

// Metrics Grids
const pdfMetricsGrid = document.getElementById('pdfMetricsGrid');
const formulaMetricsGrid = document.getElementById('formulaMetricsGrid');
const excelMetricsGrid = document.getElementById('excelMetricsGrid');
const matchMetricsGrid = document.getElementById('matchMetricsGrid');

// PDF Metrics
const statFiguresCount = document.getElementById('statFiguresCount');
const statHasAlt = document.getElementById('statHasAlt');
const statMissingAlt = document.getElementById('statMissingAlt');
const statTotalPages = document.getElementById('statTotalPages');

// Formula Metrics
const statFormulasCount = document.getElementById('statFormulasCount');
const statFormulaHasAlt = document.getElementById('statFormulaHasAlt');
const statFormulaMissingAlt = document.getElementById('statFormulaMissingAlt');
const statFormulaTotalPages = document.getElementById('statFormulaTotalPages');

// Excel Metrics
const statExcelImages = document.getElementById('statExcelImages');
const statExcelHasAlt = document.getElementById('statExcelHasAlt');
const statExcelMissingAlt = document.getElementById('statExcelMissingAlt');
const statExcelTotal = document.getElementById('statExcelTotal');

// Match Metrics
const statMatchPdfTotal = document.getElementById('statMatchPdfTotal');
const statMatchMatched = document.getElementById('statMatchMatched');
const statMatchConfidence = document.getElementById('statMatchConfidence');
const statMatchReady = document.getElementById('statMatchReady');

// Filter & Search
const filterAllBtn = document.getElementById('filterAllBtn');
const filterHasAltBtn = document.getElementById('filterHasAltBtn');
const filterMissingBtn = document.getElementById('filterMissingBtn');
const filterHasImgBtn = document.getElementById('filterHasImgBtn');
const countAll = document.getElementById('countAll');
const countHasAlt = document.getElementById('countHasAlt');
const countMissing = document.getElementById('countMissing');
const countHasImg = document.getElementById('countHasImg');
const searchInput = document.getElementById('searchInput');

// View Toggles & Containers
const viewGridBtn = document.getElementById('viewGridBtn');
const viewTableBtn = document.getElementById('viewTableBtn');
const figuresGrid = document.getElementById('figuresGrid');
const figuresTableContainer = document.getElementById('figuresTableContainer');
const figuresTableBody = document.getElementById('figuresTableBody');

// Selection Bar Elements (Above figure boxes)
const selectionBar = document.getElementById('selectionBar');
const selectAllCheckbox = document.getElementById('selectAllCheckbox');
const btnSelectAll = document.getElementById('btnSelectAll');
const btnDeselectAll = document.getElementById('btnDeselectAll');
const selectedCountBadge = document.getElementById('selectedCountBadge');
const totalSelectableBadge = document.getElementById('totalSelectableBadge');
const injectSelectedBtn = document.getElementById('injectSelectedBtn');
const injectSelectedBtnText = document.getElementById('injectSelectedBtnText');
const removeAltBtn = document.getElementById('removeAltBtn');
const removeAltBtnText = document.getElementById('removeAltBtnText');
const tableSelectAllCheckbox = document.getElementById('tableSelectAllCheckbox');
let selectedFigureIds = new Set();

// Formula Views & Pagination
const formulasGrid = document.getElementById('formulasGrid');
const formulasTableContainer = document.getElementById('formulasTableContainer');
const formulasTableBody = document.getElementById('formulasTableBody');
const formulaPaginationBar = document.getElementById('formulaPaginationBar');
const formulaPaginationInfo = document.getElementById('formulaPaginationInfo');
const formulaPageFirstBtn = document.getElementById('formulaPageFirstBtn');
const formulaPagePrevBtn = document.getElementById('formulaPagePrevBtn');
const formulaPageNextBtn = document.getElementById('formulaPageNextBtn');
const formulaPageLastBtn = document.getElementById('formulaPageLastBtn');
const formulaPageCurrentDisplay = document.getElementById('formulaPageCurrentDisplay');
const formulaPageSizeSelect = document.getElementById('formulaPageSizeSelect');
const formulaTableSelectAllCheckbox = document.getElementById('formulaTableSelectAllCheckbox');

const excelGrid = document.getElementById('excelGrid');
const excelTableContainer = document.getElementById('excelTableContainer');
const excelTableBody = document.getElementById('excelTableBody');
const matchGrid = document.getElementById('matchGrid');

// Pagination
const excelPaginationBar = document.getElementById('excelPaginationBar');
const paginationInfo = document.getElementById('paginationInfo');
const pageFirstBtn = document.getElementById('pageFirstBtn');
const pagePrevBtn = document.getElementById('pagePrevBtn');
const pageNextBtn = document.getElementById('pageNextBtn');
const pageLastBtn = document.getElementById('pageLastBtn');
const pageCurrentDisplay = document.getElementById('pageCurrentDisplay');
const pageSizeSelect = document.getElementById('pageSizeSelect');

// Modal Elements
const figureModal = document.getElementById('figureModal');
const modalCloseBtn = document.getElementById('modalCloseBtn');
const modalTitle = document.getElementById('modalTitle');
const modalSubtitle = document.getElementById('modalSubtitle');
const modalImage = document.getElementById('modalImage');
const modalDownloadLink = document.getElementById('modalDownloadLink');
const modalCopyAltBtn = document.getElementById('modalCopyAltBtn');
const modalTypeLabel = document.getElementById('modalTypeLabel');
const modalTypeDisplay = document.getElementById('modalTypeDisplay');
const modalAltLabel = document.getElementById('modalAltLabel');
const modalAltDisplay = document.getElementById('modalAltDisplay');
const modalBBoxGroup = document.getElementById('modalBBoxGroup');
const modalBBoxDisplay = document.getElementById('modalBBoxDisplay');
const modalPathGroup = document.getElementById('modalPathGroup');
const modalPathDisplay = document.getElementById('modalPathDisplay');
const modalMcidGroup = document.getElementById('modalMcidGroup');
const modalMcidDisplay = document.getElementById('modalMcidDisplay');

const modalMatchedExcelBox = document.getElementById('modalMatchedExcelBox');
const modalMatchScore = document.getElementById('modalMatchScore');
const modalExcelImage = document.getElementById('modalExcelImage');
const modalExcelRowMeta = document.getElementById('modalExcelRowMeta');
const modalExcelAltText = document.getElementById('modalExcelAltText');
const modalInjectBtn = document.getElementById('modalInjectBtn');
const modalAltTextarea = document.getElementById('modalAltTextarea');
let currentModalFigure = null;

// Toast
const toastNotification = document.getElementById('toastNotification');
const toastMessage = document.getElementById('toastMessage');
let toastTimer = null;

// ==========================================
// INITIALIZATION & EVENT LISTENERS
// ==========================================
function initEvents() {
    // PDF File Selection
    browseBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        pdfFileInput.click();
    });

    dropZone.addEventListener('click', () => {
        pdfFileInput.click();
    });

    pdfFileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
            handlePdfUpload(e.target.files[0]);
        }
    });

    // Excel File Selection
    if (browseExcelBtn) {
        browseExcelBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            excelFileInput.click();
        });
    }

    excelUploadBtn.addEventListener('click', () => {
        excelFileInput.click();
    });

    excelFileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
            handleExcelUpload(e.target.files[0]);
        }
    });

    // Sample Triggers
    if (samplePdfBtn) {
        samplePdfBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            handleLoadSamplePdf();
        });
    }

    if (sampleExcelBtn) {
        sampleExcelBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            handleLoadSampleExcel();
        });
    }

    const sampleProjectBtn = document.getElementById('sampleProjectBtn');
    if (sampleProjectBtn) {
        sampleProjectBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            handleLoadSamplePdf();
        });
    }

    // Drag & Drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            const file = e.dataTransfer.files[0];
            const lower = file.name.toLowerCase();
            if (lower.endsWith('.pdf')) {
                handlePdfUpload(file);
            } else if (lower.endsWith('.xlsx') || lower.endsWith('.xlsm')) {
                handleExcelUpload(file);
            } else {
                alert('Please upload a PDF or Excel (.xlsx) file.');
            }
        }
    });

    // Reset / Upload Another
    if (uploadAnotherBtn) {
        uploadAnotherBtn.addEventListener('click', () => {
            resultsSection.style.display = 'none';
            uploadSection.style.display = 'block';
            pdfFileInput.value = '';
            excelFileInput.value = '';
            currentSession = null;
            currentFigures = [];
            currentExcelRecords = [];
            currentMatchedFigures = [];
            selectedFigureIds.clear();
            if (selectAllCheckbox) selectAllCheckbox.checked = false;
            if (tableSelectAllCheckbox) tableSelectAllCheckbox.checked = false;
            if (typeof updateSelectionUI === 'function') updateSelectionUI();
        });
    }

    // Source Switcher Tabs
    tabPdfBtn.addEventListener('click', () => switchSourceTab('pdf'));
    if (tabFormulaBtn) {
        tabFormulaBtn.addEventListener('click', () => switchSourceTab('formula'));
    }
    tabExcelBtn.addEventListener('click', () => switchSourceTab('excel'));
    tabMatchBtn.addEventListener('click', () => switchSourceTab('match'));

    // Filter Buttons
    [filterAllBtn, filterHasAltBtn, filterMissingBtn, filterHasImgBtn].forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            activeFilter = btn.dataset.filter;
            excelPage = 1; // reset pagination on filter change
            formulaPage = 1;
            refreshActiveView();
        });
    });

    // Search Input
    searchInput.addEventListener('input', () => {
        excelPage = 1; // reset pagination on search
        formulaPage = 1;
        refreshActiveView();
    });

    // View Toggles (Grid / Table)
    viewGridBtn.addEventListener('click', () => {
        viewGridBtn.classList.add('active');
        viewTableBtn.classList.remove('active');
        currentView = 'grid';
        updateContainerVisibility();
    });

    viewTableBtn.addEventListener('click', () => {
        viewTableBtn.classList.add('active');
        viewGridBtn.classList.remove('active');
        currentView = 'table';
        updateContainerVisibility();
    });

    // Excel Pagination Event Listeners
    pageFirstBtn.addEventListener('click', () => {
        if (excelPage > 1) {
            excelPage = 1;
            renderExcel();
            scrollToTopGrid();
        }
    });

    pagePrevBtn.addEventListener('click', () => {
        if (excelPage > 1) {
            excelPage--;
            renderExcel();
            scrollToTopGrid();
        }
    });

    pageNextBtn.addEventListener('click', () => {
        const totalPages = Math.ceil(filteredExcelRecords.length / excelPageSize) || 1;
        if (excelPage < totalPages) {
            excelPage++;
            renderExcel();
            scrollToTopGrid();
        }
    });

    pageLastBtn.addEventListener('click', () => {
        const totalPages = Math.ceil(filteredExcelRecords.length / excelPageSize) || 1;
        if (excelPage < totalPages) {
            excelPage = totalPages;
            renderExcel();
            scrollToTopGrid();
        }
    });

    pageSizeSelect.addEventListener('change', (e) => {
        excelPageSize = parseInt(e.target.value, 10);
        excelPage = 1;
        renderExcel();
    });

    // Formula Pagination Event Listeners
    if (formulaPageFirstBtn) {
        formulaPageFirstBtn.addEventListener('click', () => {
            if (formulaPage > 1) {
                formulaPage = 1;
                renderFormulas();
                scrollToTopGrid();
            }
        });
    }

    if (formulaPagePrevBtn) {
        formulaPagePrevBtn.addEventListener('click', () => {
            if (formulaPage > 1) {
                formulaPage--;
                renderFormulas();
                scrollToTopGrid();
            }
        });
    }

    if (formulaPageNextBtn) {
        formulaPageNextBtn.addEventListener('click', () => {
            const totalPages = Math.ceil(filteredFormulas.length / formulaPageSize) || 1;
            if (formulaPage < totalPages) {
                formulaPage++;
                renderFormulas();
                scrollToTopGrid();
            }
        });
    }

    if (formulaPageLastBtn) {
        formulaPageLastBtn.addEventListener('click', () => {
            const totalPages = Math.ceil(filteredFormulas.length / formulaPageSize) || 1;
            if (formulaPage < totalPages) {
                formulaPage = totalPages;
                renderFormulas();
                scrollToTopGrid();
            }
        });
    }

    if (formulaPageSizeSelect) {
        formulaPageSizeSelect.addEventListener('change', (e) => {
            formulaPageSize = parseInt(e.target.value, 10);
            formulaPage = 1;
            renderFormulas();
        });
    }

    // Downloads
    if (downloadUnselectedBtn) {
        downloadUnselectedBtn.addEventListener('click', () => {
            if (!currentSession || !currentSession.session_id) return;
            if (currentSourceTab === 'excel') {
                window.location.href = `/api/download-excel-zip/${currentSession.session_id}`;
            } else if (currentSourceTab === 'formula') {
                window.location.href = `/api/download-formulas-zip/${currentSession.session_id}`;
            } else {
                downloadUnselectedFigures();
            }
        });
    }

    if (downloadFormulasZipBtn) {
        downloadFormulasZipBtn.addEventListener('click', () => {
            if (currentSession && currentSession.session_id) {
                window.location.href = `/api/download-formulas-zip/${currentSession.session_id}`;
            }
        });
    }

    downloadExcelZipBtn.addEventListener('click', () => {
        if (currentSession && currentSession.session_id) {
            window.location.href = `/api/download-excel-zip/${currentSession.session_id}`;
        }
    });

    exportJsonBtn.addEventListener('click', () => {
        if (currentSession && currentSession.session_id) {
            window.location.href = `/api/download-json/${currentSession.session_id}`;
        }
    });

    // Modal Close
    modalCloseBtn.addEventListener('click', () => {
        figureModal.style.display = 'none';
    });

    window.addEventListener('click', (e) => {
        if (e.target === figureModal) {
            figureModal.style.display = 'none';
        }
    });

    // Modal Copy Alt
    modalCopyAltBtn.addEventListener('click', () => {
        const text = (modalAltTextarea && modalAltTextarea.value) ? modalAltTextarea.value : (modalAltDisplay ? modalAltDisplay.textContent : '');
        if (text && text !== 'No Alt text present') {
            copyToClipboard(text, 'Authoritative ALT text copied!');
        }
    });

    // Alt Text Injection Button (Batch)
    if (injectAltBtn) {
        injectAltBtn.addEventListener('click', () => {
            handleBatchInjectAlt();
        });
    }

    // Modal Single Figure / Formula Alt Injection
    if (modalInjectBtn) {
        modalInjectBtn.addEventListener('click', () => {
            if (currentModalFigure) {
                const altText = modalAltTextarea ? modalAltTextarea.value.trim() : '';
                window.injectAltForFigure(currentModalFigure.figure_id, altText);
            } else if (currentModalFormula) {
                const altText = modalAltTextarea ? modalAltTextarea.value.trim() : '';
                window.injectAltForFormula(currentModalFormula.formula_id, altText);
            }
        });
    }

    // Selection Bar Listeners (Above figure/formula boxes)
    if (selectAllCheckbox) {
        selectAllCheckbox.addEventListener('change', (e) => {
            if (currentSourceTab === 'formula') {
                if (e.target.checked) window.selectAllFormulas();
                else window.deselectAllFormulas();
            } else {
                if (e.target.checked) window.selectAllFigures();
                else window.deselectAllFigures();
            }
        });
    }

    if (tableSelectAllCheckbox) {
        tableSelectAllCheckbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                window.selectAllFigures();
            } else {
                window.deselectAllFigures();
            }
        });
    }

    if (formulaTableSelectAllCheckbox) {
        formulaTableSelectAllCheckbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                window.selectAllFormulas();
            } else {
                window.deselectAllFormulas();
            }
        });
    }

    if (btnSelectAll) {
        btnSelectAll.addEventListener('click', () => {
            if (currentSourceTab === 'formula') {
                window.selectAllFormulas();
            } else {
                window.selectAllFigures();
            }
        });
    }

    if (btnDeselectAll) {
        btnDeselectAll.addEventListener('click', () => {
            if (currentSourceTab === 'formula') {
                window.deselectAllFormulas();
            } else {
                window.deselectAllFigures();
            }
        });
    }

    if (injectSelectedBtn) {
        injectSelectedBtn.addEventListener('click', () => {
            handleInjectSelected();
        });
    }

    if (removeAltBtn) {
        removeAltBtn.addEventListener('click', () => {
            handleRemoveAlt();
        });
    }
}

function scrollToTopGrid() {
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ==========================================
// TOAST NOTIFICATION
// ==========================================
function showToast(msg) {
    if (toastTimer) clearTimeout(toastTimer);
    toastMessage.textContent = msg;
    toastNotification.classList.add('show');
    toastTimer = setTimeout(() => {
        toastNotification.classList.remove('show');
    }, 3000);
}

function copyToClipboard(text, successMsg = 'Copied to clipboard!') {
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
        showToast(successMsg);
    }).catch(() => {
        // Fallback
        const ta = document.createElement('textarea');
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        showToast(successMsg);
    });
}

// ==========================================
// UPLOAD HANDLERS
// ==========================================
async function handlePdfUpload(file) {
    showLoading('Uploading and parsing PDF...', 'Traversing StructTreeRoot and extracting /Figure tags...');
    const formData = new FormData();
    formData.append('file', file);
    if (currentSession && currentSession.session_id) {
        formData.append('session_id', currentSession.session_id);
    }

    try {
        const res = await fetch('/api/upload-pdf', {
            method: 'POST',
            body: formData
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'PDF upload failed');
        }

        const data = await res.json();
        onPdfLoaded(data);
    } catch (err) {
        alert('Error: ' + err.message);
        hideLoading();
    }
}

async function handleLoadSamplePdf() {
    showLoading('Loading Chapter 15 Sample PDF...', 'Extracting 88 /Figure tags and marked content coordinates...');
    try {
        const res = await fetch('/api/load-sample', { method: 'POST' });
        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Failed to load sample PDF');
        }
        const data = await res.json();
        onPdfLoaded(data);
    } catch (err) {
        alert('Error loading sample: ' + err.message);
        hideLoading();
    }
}

async function handleExcelUpload(file) {
    const sid = currentSession ? currentSession.session_id : '';
    showLoading('Uploading Excel ALT Manifest...', 'Extracting drawings, filenames, and authoritative ALT text...');
    const formData = new FormData();
    formData.append('file', file);
    if (sid) formData.append('session_id', sid);

    try {
        const res = await fetch('/api/upload-excel', {
            method: 'POST',
            body: formData
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Excel upload failed');
        }

        const data = await res.json();
        onExcelLoaded(data);
    } catch (err) {
        alert('Error loading Excel: ' + err.message);
        hideLoading();
    }
}

async function handleLoadSampleExcel() {
    const sid = currentSession ? currentSession.session_id : '';
    showLoading('Loading Plesha Chapter 15 Excel Manifest...', 'Parsing embedded drawings and authoritative ALT texts...');

    try {
        const url = sid ? `/api/load-sample-excel?session_id=${sid}` : '/api/load-sample-excel';
        const res = await fetch(url, { method: 'POST' });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Failed to load sample Excel');
        }

        const data = await res.json();
        onExcelLoaded(data);
    } catch (err) {
        alert('Error loading sample Excel: ' + err.message);
        hideLoading();
    }
}

// ==========================================
// DATA LOADED CALLBACKS
// ==========================================
function onPdfLoaded(data) {
    currentSession = data;
    currentFigures = data.figures || [];
    currentFormulas = data.formulas || [];

    hideLoading();
    resultsSection.style.display = 'block';

    // Update PDF Figure Metrics
    statFiguresCount.textContent = data.figures_count;
    statHasAlt.textContent = data.has_alt_count;
    statMissingAlt.textContent = data.missing_alt_count;
    statTotalPages.textContent = data.total_pages;

    tabPdfBadge.textContent = data.figures_count;

    // Update PDF Formula Metrics
    if (statFormulasCount) statFormulasCount.textContent = (data.formulas_count || currentFormulas.length).toLocaleString();
    if (statFormulaHasAlt) statFormulaHasAlt.textContent = (data.has_formula_alt_count || currentFormulas.filter(f => f.has_alt).length).toLocaleString();
    if (statFormulaMissingAlt) statFormulaMissingAlt.textContent = (data.missing_formula_alt_count || (currentFormulas.length - (data.has_formula_alt_count || 0))).toLocaleString();
    if (statFormulaTotalPages) statFormulaTotalPages.textContent = data.total_pages;
    if (tabFormulaBadge) tabFormulaBadge.textContent = currentFormulas.length.toLocaleString();

    // Injected PDF download visibility
    if (downloadInjectedPdfBtn) {
        if (data.has_injected_pdf) {
            downloadInjectedPdfBtn.href = `/api/download-injected-pdf/${data.session_id}`;
            downloadInjectedPdfBtn.style.display = 'inline-flex';
        } else {
            downloadInjectedPdfBtn.style.display = 'none';
        }
    }
    if (injectAltBtn) {
        injectAltBtn.style.display = 'inline-flex';
        injectAltBtn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>
            <span id="injectAltBtnText">⚡ Inject All Alt into PDF</span>
        `;
        injectAltBtn.disabled = false;
    }

    // Check if Excel already loaded in session
    if (data.excel_records && data.excel_records.length > 0) {
        currentExcelRecords = data.excel_records;
        tabExcelBadge.textContent = currentExcelRecords.length.toLocaleString();
    }

    // Pre-select all figures that have ALT text
    selectedFigureIds = new Set(
        currentFigures
            .filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text)
            .map(f => f.figure_id)
    );

    // Pre-select formulas
    selectedFormulaIds = new Set(
        currentFormulas
            .filter(f => f.alt_text || f.actual_text)
            .map(f => f.formula_id)
    );

    updateTabVisibility();
    switchSourceTab('pdf');
}

function onExcelLoaded(data) {
    if (!currentSession) {
        currentSession = data;
    } else {
        // Merge session
        currentSession.excel_filename = data.excel_filename;
        currentSession.excel_total_records = data.excel_total_records;
        currentSession.excel_records = data.excel_records;
        currentSession.excel_images_count = data.excel_images_count;
        currentSession.excel_has_alt_count = data.excel_has_alt_count;
        currentSession.excel_missing_alt_count = data.excel_missing_alt_count;
    }

    currentExcelRecords = data.excel_records || [];
    if (data.matched_figures && data.matched_figures.length > 0) {
        currentFigures = data.matched_figures;
        currentMatchedFigures = data.matched_figures.filter(f => f.excel_match);
        // Refresh selection with newly matched figures
        selectedFigureIds = new Set(
            currentFigures
                .filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text)
                .map(f => f.figure_id)
        );
    }
    if (data.matched_formulas && data.matched_formulas.length > 0) {
        currentFormulas = data.matched_formulas;
        // Refresh formula selection with newly matched formulas
        selectedFormulaIds = new Set(
            currentFormulas
                .filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text || f.actual_text)
                .map(f => f.formula_id)
        );
    }

    hideLoading();
    resultsSection.style.display = 'block';

    // Update Excel Metrics
    statExcelImages.textContent = (data.excel_images_count || currentExcelRecords.length).toLocaleString();
    statExcelHasAlt.textContent = (data.excel_has_alt_count || currentExcelRecords.filter(r => r.has_alt).length).toLocaleString();
    statExcelMissingAlt.textContent = (data.excel_missing_alt_count || 0).toLocaleString();
    statExcelTotal.textContent = (data.excel_total_records || currentExcelRecords.length).toLocaleString();

    tabExcelBadge.textContent = currentExcelRecords.length.toLocaleString();

    if (sampleExcelBtn) {
        sampleExcelBtn.innerHTML = '✓ Excel Loaded';
        sampleExcelBtn.classList.add('active');
    }

    updateTabVisibility();

    // Directly switch to Excel tab to display all Excel images and alt text below!
    switchSourceTab('excel');
    showToast(`Loaded ${currentExcelRecords.length.toLocaleString()} Excel records with drawings & ALT text!`);
}

function updateTabVisibility() {
    const hasPdf = currentFigures && currentFigures.length > 0;
    const hasFormulas = currentFormulas && currentFormulas.length > 0;
    const hasExcel = currentExcelRecords && currentExcelRecords.length > 0;

    tabPdfBtn.style.display = hasPdf ? 'inline-flex' : 'none';
    if (tabFormulaBtn) tabFormulaBtn.style.display = hasFormulas ? 'inline-flex' : 'none';
    tabExcelBtn.style.display = hasExcel ? 'inline-flex' : 'none';

    if (hasPdf && hasExcel) {
        tabMatchBtn.style.display = 'inline-flex';
        const matchedFigures = currentFigures.filter(f => f.excel_match);
        const matchedCount = matchedFigures.length;
        tabMatchBadge.textContent = matchedCount;
        statMatchPdfTotal.textContent = currentFigures.length;
        statMatchMatched.textContent = matchedCount;
        statMatchReady.textContent = matchedCount;

        if (matchedCount > 0) {
            const totalConfidence = matchedFigures.reduce((sum, f) => {
                const conf = (typeof f.confidence === 'number' && !isNaN(f.confidence)) ? f.confidence : (f.excel_match ? 0.95 : 0);
                return sum + conf;
            }, 0);
            const avgConfPct = Math.round((totalConfidence / matchedCount) * 100);
            statMatchConfidence.textContent = `${avgConfPct}%`;
        } else {
            statMatchConfidence.textContent = '0%';
        }
    } else {
        tabMatchBtn.style.display = 'none';
        statMatchConfidence.textContent = '0%';
    }
}

// ==========================================
// SOURCE TAB SWITCHING
// ==========================================
function switchSourceTab(tab) {
    currentSourceTab = tab;

    // Update Tab UI buttons
    tabPdfBtn.classList.toggle('active', tab === 'pdf');
    if (tabFormulaBtn) tabFormulaBtn.classList.toggle('active', tab === 'formula');
    tabExcelBtn.classList.toggle('active', tab === 'excel');
    tabMatchBtn.classList.toggle('active', tab === 'match');

    // Toggle Metrics Grids
    pdfMetricsGrid.style.display = tab === 'pdf' ? 'grid' : 'none';
    if (formulaMetricsGrid) formulaMetricsGrid.style.display = tab === 'formula' ? 'grid' : 'none';
    excelMetricsGrid.style.display = tab === 'excel' ? 'grid' : 'none';
    matchMetricsGrid.style.display = tab === 'match' ? 'grid' : 'none';

    // Show/Hide Download Buttons & Selection Bar
    if (tab === 'excel') {
        docBadge.textContent = 'Excel ALT Manifest';
        docBadge.style.color = '#34d399';
        docBadge.style.borderColor = 'rgba(16, 185, 129, 0.3)';
        docBadge.style.background = 'rgba(16, 185, 129, 0.12)';
        docFilename.textContent = currentSession ? (currentSession.excel_filename || 'Manifest.xlsx') : 'Excel Manifest';
        docMeta.textContent = `${(currentExcelRecords.length).toLocaleString()} Drawing Rows • Authoritative ALT Text Gallery`;
        if (downloadUnselectedLabel) downloadUnselectedLabel.textContent = 'Download Excel Images (ZIP)';
        if (downloadFormulasZipBtn) downloadFormulasZipBtn.style.display = 'none';
        downloadExcelZipBtn.style.display = 'none';
        if (injectAltBtn) injectAltBtn.style.display = 'none';
        if (downloadInjectedPdfBtn) downloadInjectedPdfBtn.style.display = 'none';
        if (selectionBar) selectionBar.style.display = 'none';
        filterHasImgBtn.style.display = 'inline-flex';
        searchInput.placeholder = 'Search by Row #, Filename (e.g. epub_img_161.png), or ALT keyword...';
    } else if (tab === 'formula') {
        docBadge.textContent = 'PDF Math Formulas';
        docBadge.style.color = 'var(--primary)';
        docBadge.style.borderColor = 'rgba(56, 189, 248, 0.25)';
        docBadge.style.background = 'rgba(56, 189, 248, 0.12)';
        docFilename.textContent = currentSession ? (currentSession.filename || 'Document.pdf') : 'PDF Formulas';
        docMeta.textContent = `${(currentFormulas.length).toLocaleString()} MathType / Math Formula Tags • StructTreeRoot Engine`;
        if (downloadUnselectedLabel) downloadUnselectedLabel.textContent = 'Download Formulas (ZIP)';
        if (downloadFormulasZipBtn) downloadFormulasZipBtn.style.display = 'inline-flex';
        downloadExcelZipBtn.style.display = 'none';
        if (injectAltBtn) {
            injectAltBtn.style.display = 'inline-flex';
            injectAltBtn.innerHTML = `
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>
                <span id="injectAltBtnText">⚡ Inject All Formula Alt</span>
            `;
        }
        if (downloadInjectedPdfBtn) {
            downloadInjectedPdfBtn.style.display = (currentSession && currentSession.has_injected_pdf) ? 'inline-flex' : 'none';
        }
        if (selectionBar) selectionBar.style.display = 'flex';
        filterHasImgBtn.style.display = 'none';
        searchInput.placeholder = 'Search by Formula #, Page #, MCID, or Alt text...';
    } else if (tab === 'pdf') {
        docBadge.textContent = 'Accessible PDF';
        docBadge.style.color = 'var(--primary)';
        docBadge.style.borderColor = 'rgba(56, 189, 248, 0.25)';
        docBadge.style.background = 'rgba(56, 189, 248, 0.12)';
        docFilename.textContent = currentSession ? (currentSession.filename || 'Document.pdf') : 'PDF Figures';
        docMeta.textContent = `${currentSession ? currentSession.total_pages : 0} Pages • StructTreeRoot Tagged Engine`;
        if (downloadUnselectedLabel) downloadUnselectedLabel.textContent = 'Download Unselected Figures (Excel)';
        if (downloadFormulasZipBtn) downloadFormulasZipBtn.style.display = 'none';
        downloadExcelZipBtn.style.display = currentExcelRecords.length > 0 ? 'inline-flex' : 'none';
        if (injectAltBtn) {
            injectAltBtn.style.display = 'inline-flex';
            injectAltBtn.innerHTML = `
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>
                <span id="injectAltBtnText">⚡ Inject All Alt into PDF</span>
            `;
        }
        if (downloadInjectedPdfBtn) {
            downloadInjectedPdfBtn.style.display = (currentSession && currentSession.has_injected_pdf) ? 'inline-flex' : 'none';
        }
        if (selectionBar) selectionBar.style.display = 'flex';
        filterHasImgBtn.style.display = 'none';
        searchInput.placeholder = 'Search by Page #, MCID, or ID...';
    } else if (tab === 'match') {
        docBadge.textContent = 'Side-by-Side Match';
        docBadge.style.color = 'var(--accent-purple)';
        docBadge.style.borderColor = 'rgba(168, 85, 247, 0.3)';
        docBadge.style.background = 'rgba(168, 85, 247, 0.12)';
        docFilename.textContent = `${currentSession ? currentSession.filename : 'PDF'} ⟷ ${currentSession ? currentSession.excel_filename : 'Excel'}`;
        docMeta.textContent = 'Automated Visual Alignment of PDF Figures to Authoritative Excel ALT';
        if (downloadUnselectedLabel) downloadUnselectedLabel.textContent = 'Download Figures (ZIP)';
        if (downloadFormulasZipBtn) downloadFormulasZipBtn.style.display = 'none';
        if (injectAltBtn) {
            injectAltBtn.style.display = 'inline-flex';
            injectAltBtn.innerHTML = `
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>
                <span id="injectAltBtnText">⚡ Inject All Alt into PDF</span>
            `;
        }
        if (downloadInjectedPdfBtn) {
            downloadInjectedPdfBtn.style.display = (currentSession && currentSession.has_injected_pdf) ? 'inline-flex' : 'none';
        }
        if (selectionBar) selectionBar.style.display = 'flex';
        filterHasImgBtn.style.display = 'none';
        searchInput.placeholder = 'Search matched figures by page, MCID, or ALT...';
    }

    // Refresh view
    refreshActiveView();
}

function refreshActiveView() {
    updateFilterCounts();
    updateContainerVisibility();

    if (currentSourceTab === 'pdf') {
        renderPdfFigures();
    } else if (currentSourceTab === 'formula') {
        renderFormulas();
    } else if (currentSourceTab === 'excel') {
        renderExcel();
    } else if (currentSourceTab === 'match') {
        renderMatched();
    }
}

function updateContainerVisibility() {
    const isGrid = currentView === 'grid';

    // Hide all first
    figuresGrid.style.display = 'none';
    figuresTableContainer.style.display = 'none';
    if (formulasGrid) formulasGrid.style.display = 'none';
    if (formulasTableContainer) formulasTableContainer.style.display = 'none';
    if (formulaPaginationBar) formulaPaginationBar.style.display = 'none';
    excelGrid.style.display = 'none';
    excelTableContainer.style.display = 'none';
    matchGrid.style.display = 'none';
    excelPaginationBar.style.display = 'none';

    if (currentSourceTab === 'pdf') {
        if (isGrid) figuresGrid.style.display = 'grid';
        else figuresTableContainer.style.display = 'block';
    } else if (currentSourceTab === 'formula') {
        if (isGrid) {
            if (formulasGrid) formulasGrid.style.display = 'grid';
        } else {
            if (formulasTableContainer) formulasTableContainer.style.display = 'block';
        }
        if (formulaPaginationBar) formulaPaginationBar.style.display = 'flex';
    } else if (currentSourceTab === 'excel') {
        if (isGrid) excelGrid.style.display = 'grid';
        else excelTableContainer.style.display = 'block';
        excelPaginationBar.style.display = 'flex';
    } else if (currentSourceTab === 'match') {
        matchGrid.style.display = 'flex';
    }
}

// ==========================================
// FILTER COUNTS CALCULATION
// ==========================================
function updateFilterCounts() {
    if (currentSourceTab === 'pdf') {
        const total = currentFigures.length;
        const hasAlt = currentFigures.filter(f => f.has_alt).length;
        const missing = total - hasAlt;

        countAll.textContent = total;
        countHasAlt.textContent = hasAlt;
        countMissing.textContent = missing;
    } else if (currentSourceTab === 'formula') {
        const total = currentFormulas.length;
        const hasAlt = currentFormulas.filter(f => f.has_alt).length;
        const missing = total - hasAlt;

        countAll.textContent = total.toLocaleString();
        countHasAlt.textContent = hasAlt.toLocaleString();
        countMissing.textContent = missing.toLocaleString();
    } else if (currentSourceTab === 'excel') {
        const total = currentExcelRecords.length;
        const hasAlt = currentExcelRecords.filter(r => r.has_alt).length;
        const missing = total - hasAlt;
        const withImg = currentExcelRecords.filter(r => r.has_image).length;

        countAll.textContent = total.toLocaleString();
        countHasAlt.textContent = hasAlt.toLocaleString();
        countMissing.textContent = missing.toLocaleString();
        countHasImg.textContent = withImg.toLocaleString();
    } else if (currentSourceTab === 'match') {
        const total = currentFigures.length;
        const matched = currentFigures.filter(f => f.excel_match).length;
        countAll.textContent = total;
        countHasAlt.textContent = matched;
        countMissing.textContent = total - matched;
    }
}

// ==========================================
// PDF FIGURES RENDERING
// ==========================================
function renderPdfFigures() {
    const q = searchInput.value.toLowerCase().trim();

    const filtered = currentFigures.filter(fig => {
        if (activeFilter === 'missing' && fig.has_alt) return false;
        if (activeFilter === 'has_alt' && !fig.has_alt) return false;

        if (q) {
            const pageMatch = String(fig.page_number).includes(q);
            const idMatch = String(fig.figure_id).includes(q);
            const mcidMatch = (fig.mcids || []).some(m => String(m).includes(q));
            const altMatch = (fig.alt_text || '').toLowerCase().includes(q);
            const exAltMatch = (fig.excel_match && fig.excel_match.alt_text || '').toLowerCase().includes(q);
            return pageMatch || idMatch || mcidMatch || altMatch || exAltMatch;
        }
        return true;
    });

    // Render Grid
    figuresGrid.innerHTML = '';
    if (filtered.length === 0) {
        figuresGrid.innerHTML = `
            <div style="grid-column: 1 / -1; text-align: center; padding: 60px 20px; color: var(--text-dim);">
                <h3>No PDF figures match the current filter or search criteria</h3>
            </div>
        `;
    } else {
        filtered.forEach(fig => {
            const isSelected = selectedFigureIds.has(fig.figure_id);
            const card = document.createElement('div');
            card.className = `figure-card ${isSelected ? 'selected' : ''}`;
            card.setAttribute('data-fig-id', fig.figure_id);
            card.addEventListener('click', () => openPdfModal(fig));

            const mcidText = fig.mcids && fig.mcids.length > 0 ? `MCID ${fig.mcids.join(',')}` : 'Vector/Image';
            const ex = fig.excel_match;
            const effectiveAlt = (ex && ex.alt_text) || fig.alt_text || '';
            const hasEffectiveAlt = Boolean(effectiveAlt);
            const isInjected = fig.status_label === 'Injected';
            const statusClass = isInjected ? 'present' : (ex ? 'present' : (fig.has_alt ? 'present' : 'missing'));
            const statusLabel = isInjected ? '✓ Alt Injected' : (ex ? `✓ Alt Linked (Row ${ex.row})` : (fig.has_alt ? 'Has /Alt' : 'Missing /Alt'));

            let imagesHtml = '';
            if (ex && ex.image_url) {
                const confVal = (typeof fig.confidence === 'number' && !isNaN(fig.confidence)) ? fig.confidence : 0.95;
                const confPct = Math.round(confVal * 100);
                imagesHtml = `
                    <div class="card-dual-image-grid">
                        <div class="figure-image-wrapper">
                            <span class="page-chip">Page ${fig.page_number || '?'}</span>
                            <span class="mcid-chip">${mcidText}</span>
                            <img src="${fig.image_url || ''}" alt="Figure ${fig.figure_id}" loading="lazy">
                            <span class="img-type-badge">PDF Structure Crop</span>
                        </div>
                        <div class="figure-image-wrapper excel-preview-wrapper">
                            <span class="row-chip">Excel Row ${ex.row} (Sr. ${ex.sr_no})</span>
                            <img src="${ex.image_url}" alt="Excel Drawing" loading="lazy">
                            <span class="img-type-badge excel-type-badge">Excel Manifest Image</span>
                        </div>
                    </div>
                `;
            } else {
                imagesHtml = `
                    <div class="figure-image-wrapper">
                        <span class="page-chip">Page ${fig.page_number || '?'}</span>
                        <span class="mcid-chip">${mcidText}</span>
                        <img src="${fig.image_url || ''}" alt="Figure ${fig.figure_id}" loading="lazy">
                    </div>
                `;
            }

            card.innerHTML = `
                <!-- Select Box Option Bar (Above each box) -->
                <div class="card-select-bar">
                    <label class="card-select-control" onclick="event.stopPropagation();">
                        <input type="checkbox" class="card-checkbox" data-fig-id="${fig.figure_id}" ${isSelected ? 'checked' : ''} onchange="window.toggleFigureSelection(${fig.figure_id}, this.checked)">
                        <span class="card-select-text">Select Box for Injection</span>
                    </label>
                    <button class="btn-card-select-toggle ${isSelected ? 'selected' : ''}" onclick="event.stopPropagation(); window.toggleFigureSelection(${fig.figure_id})">
                        ${isSelected ? '✓ Selected' : '+ Select'}
                    </button>
                </div>
                ${imagesHtml}
                <div class="figure-content">
                    <div class="figure-title-row">
                        <span class="figure-id">Figure ${fig.figure_id}</span>
                        <span class="status-tag ${statusClass}">${statusLabel}</span>
                    </div>
                    <div class="figure-meta-row">
                        <span>BBox: ${fig.bbox_width} × ${fig.bbox_height} pt</span>
                        <span style="color:${ex ? '#34d399' : 'inherit'}; font-weight:600;">${ex ? ex.filename : `XObjs: ${fig.underlying_xobjects_count}`}</span>
                    </div>
                    <div class="figure-alt-preview ${hasEffectiveAlt ? 'has-alt' : 'empty'}">
                        <div class="alt-label-bar">
                            <span>${isInjected ? '✓ INJECTED STRUCTTREE ALT:' : (ex ? 'AUTHORITATIVE ALT TEXT FROM EXCEL:' : 'CURRENT PDF /ALT:')}</span>
                            <div style="display:flex; gap:6px;">
                                ${hasEffectiveAlt ? `<button class="btn-copy-alt-mini" onclick="event.stopPropagation(); window.copyFigureAlt(${fig.figure_id})">Copy</button>` : ''}
                                ${hasEffectiveAlt ? `<button class="btn-inject-mini" onclick="event.stopPropagation(); window.injectAltForFigure(${fig.figure_id})" title="Inject into PDF StructTree">⚡ Inject</button>` : ''}
                            </div>
                        </div>
                        <div class="alt-body-text">${escapeHtml(effectiveAlt || 'No /Alt accessibility text defined in StructTree.')}</div>
                    </div>
                    <div class="card-footer-row">
                        <span class="footer-hint">${ex ? `Row ${ex.row} • Sr. ${ex.sr_no}` : `Page ${fig.page_number}`}</span>
                        <div style="display:flex; gap:6px;">
                            ${hasEffectiveAlt ? `
                            <button class="btn btn-inject btn-sm" onclick="event.stopPropagation(); window.injectAltForFigure(${fig.figure_id})" style="padding: 5px 10px; font-size: 0.78rem;">
                                ⚡ Inject Alt
                            </button>` : ''}
                            <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); window.openPdfModalById(${fig.figure_id})">
                                Inspect Details
                            </button>
                        </div>
                    </div>
                </div>
            `;
            figuresGrid.appendChild(card);
        });
    }

    // Render Table
    figuresTableBody.innerHTML = '';
    filtered.forEach(fig => {
        const isSelected = selectedFigureIds.has(fig.figure_id);
        const tr = document.createElement('tr');
        const isMissing = !fig.has_alt;
        const isInjected = fig.status_label === 'Injected';
        const statusClass = isInjected ? 'present' : (isMissing ? 'missing' : 'present');
        const statusLabel = isInjected ? '✓ Alt Injected' : (isMissing ? 'Missing /Alt' : 'Has /Alt');
        const mcidText = fig.mcids && fig.mcids.length > 0 ? fig.mcids.join(', ') : 'None';
        const ex = fig.excel_match;
        const effectiveAlt = (ex && ex.alt_text) || fig.alt_text || '';
        const altSummary = effectiveAlt || 'None';

        tr.innerHTML = `
            <td style="text-align: center;">
                <input type="checkbox" class="table-card-checkbox" data-fig-id="${fig.figure_id}" ${isSelected ? 'checked' : ''} onchange="event.stopPropagation(); window.toggleFigureSelection(${fig.figure_id}, this.checked)">
            </td>
            <td><strong>#${fig.figure_id}</strong></td>
            <td>
                <div class="table-thumb">
                    <img src="${fig.image_url || ''}" alt="Fig ${fig.figure_id}" loading="lazy">
                </div>
            </td>
            <td>Page ${fig.page_number || '?'}</td>
            <td><code>${mcidText}</code></td>
            <td>${fig.bbox_width} × ${fig.bbox_height}</td>
            <td><span class="status-tag ${statusClass}">${statusLabel}</span></td>
            <td><span style="font-size:0.8rem; color:${ex ? '#34d399' : 'var(--text-muted)'};">${escapeHtml(altSummary.substring(0, 70))}...</span></td>
            <td>
                <div style="display:flex; gap:6px;">
                    ${effectiveAlt ? `
                    <button class="btn btn-inject btn-sm" onclick="event.stopPropagation(); window.injectAltForFigure(${fig.figure_id})" style="padding: 4px 8px; font-size: 0.75rem;">
                        ⚡ Inject
                    </button>` : ''}
                    <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); window.openPdfModalById(${fig.figure_id})">
                        Inspect
                    </button>
                </div>
            </td>
        `;
        tr.addEventListener('click', () => openPdfModal(fig));
        figuresTableBody.appendChild(tr);
    });

    updateSelectionUI();
}

// ==========================================
// PDF FORMULAS RENDERING & PAGINATION
// ==========================================
function renderFormulas() {
    const q = searchInput.value.toLowerCase().trim();

    filteredFormulas = currentFormulas.filter(formula => {
        const hasAlt = formula.has_alt || Boolean(formula.excel_match && formula.excel_match.alt_text);
        if (activeFilter === 'missing' && hasAlt) return false;
        if (activeFilter === 'has_alt' && !hasAlt) return false;

        if (q) {
            const pageMatch = String(formula.page_number).includes(q);
            const idMatch = String(formula.formula_id).includes(q);
            const mcidMatch = (formula.mcids || []).some(m => String(m).includes(q));
            const altMatch = (formula.alt_text || '').toLowerCase().includes(q);
            const exAltMatch = (formula.excel_match && formula.excel_match.alt_text || '').toLowerCase().includes(q);
            const actualMatch = (formula.actual_text || '').toLowerCase().includes(q);
            const typeMatch = (formula.formula_type || '').toLowerCase().includes(q);
            return pageMatch || idMatch || mcidMatch || altMatch || exAltMatch || actualMatch || typeMatch;
        }
        return true;
    });

    const totalFiltered = filteredFormulas.length;
    const totalPages = Math.ceil(totalFiltered / formulaPageSize) || 1;
    if (formulaPage > totalPages) formulaPage = totalPages;
    if (formulaPage < 1) formulaPage = 1;

    const startIdx = (formulaPage - 1) * formulaPageSize;
    const endIdx = Math.min(startIdx + formulaPageSize, totalFiltered);
    const pageRecords = filteredFormulas.slice(startIdx, endIdx);

    // Update Pagination UI
    if (formulaPaginationInfo) {
        formulaPaginationInfo.textContent = totalFiltered > 0
            ? `Showing ${(startIdx + 1).toLocaleString()} - ${endIdx.toLocaleString()} of ${totalFiltered.toLocaleString()} Math Formulas`
            : 'No formulas found';
    }
    if (formulaPageCurrentDisplay) {
        formulaPageCurrentDisplay.textContent = `Page ${formulaPage.toLocaleString()} of ${totalPages.toLocaleString()}`;
    }

    if (formulaPageFirstBtn) formulaPageFirstBtn.disabled = formulaPage <= 1;
    if (formulaPagePrevBtn) formulaPagePrevBtn.disabled = formulaPage <= 1;
    if (formulaPageNextBtn) formulaPageNextBtn.disabled = formulaPage >= totalPages;
    if (formulaPageLastBtn) formulaPageLastBtn.disabled = formulaPage >= totalPages;

    // Render Grid
    if (formulasGrid) {
        formulasGrid.innerHTML = '';
        if (pageRecords.length === 0) {
            formulasGrid.innerHTML = `
                <div style="grid-column: 1 / -1; text-align: center; padding: 60px 20px; color: var(--text-dim);">
                    <h3>No PDF formulas match the current filter or search criteria</h3>
                </div>
            `;
        } else {
            pageRecords.forEach(formula => {
                const isSelected = selectedFormulaIds.has(formula.formula_id);
                const card = document.createElement('div');
                card.className = `figure-card ${isSelected ? 'selected' : ''}`;
                card.setAttribute('data-formula-id', formula.formula_id);
                card.addEventListener('click', () => openFormulaModal(formula));

                const mcidText = formula.mcids && formula.mcids.length > 0 ? `MCID ${formula.mcids.join(',')}` : 'Formula Crop';
                const ex = formula.excel_match;
                const effectiveAlt = (ex && ex.alt_text) || formula.alt_text || formula.actual_text || '';
                const hasEffectiveAlt = Boolean(effectiveAlt);
                const isInjected = formula.status_label === 'Injected';
                const statusClass = isInjected ? 'present' : (ex ? 'present' : (formula.has_alt ? 'present' : 'missing'));
                const statusLabel = isInjected ? '✓ Alt Injected' : (ex ? `✓ Alt Linked (Row ${ex.row})` : (formula.has_alt ? 'Has /Alt' : 'Missing /Alt'));

                let imagesHtml = '';
                if (ex && ex.image_url) {
                    const confVal = (typeof formula.confidence === 'number' && !isNaN(formula.confidence)) ? formula.confidence : 0.95;
                    const confPct = Math.round(confVal * 100);
                    imagesHtml = `
                        <div class="card-dual-image-grid">
                            <div class="figure-image-wrapper">
                                <span class="page-chip">Page ${formula.page_number || '?'}</span>
                                <span class="mcid-chip">${mcidText}</span>
                                <img src="${formula.image_url || ''}" alt="Formula ${formula.formula_id}" loading="lazy">
                                <span class="img-type-badge">PDF Formula Crop</span>
                            </div>
                            <div class="figure-image-wrapper excel-preview-wrapper">
                                <span class="row-chip">Excel Row ${ex.row} (Sr. ${ex.sr_no})</span>
                                <img src="${ex.image_url}" alt="Excel Math Drawing" loading="lazy">
                                <span class="img-type-badge excel-type-badge">Excel Manifest Image</span>
                            </div>
                        </div>
                    `;
                } else {
                    imagesHtml = `
                        <div class="figure-image-wrapper">
                            <span class="page-chip">Page ${formula.page_number || '?'}</span>
                            <span class="mcid-chip">${mcidText}</span>
                            <img src="${formula.image_url || ''}" alt="Formula ${formula.formula_id}" loading="lazy">
                            <span class="img-type-badge">Tag /Formula Crop</span>
                        </div>
                    `;
                }

                card.innerHTML = `
                    <!-- Select Box Option Bar (Above each formula box) -->
                    <div class="card-select-bar">
                        <label class="card-select-control" onclick="event.stopPropagation();">
                            <input type="checkbox" class="card-checkbox" data-formula-id="${formula.formula_id}" ${isSelected ? 'checked' : ''} onchange="window.toggleFormulaSelection(${formula.formula_id}, this.checked)">
                            <span class="card-select-text">Select Formula for Injection</span>
                        </label>
                        <button class="btn-card-select-toggle ${isSelected ? 'selected' : ''}" onclick="event.stopPropagation(); window.toggleFormulaSelection(${formula.formula_id})">
                            ${isSelected ? '✓ Selected' : '+ Select'}
                        </button>
                    </div>
                    ${imagesHtml}
                    <div class="figure-content">
                        <div class="figure-title-row">
                            <span class="figure-id">Formula #${formula.formula_id}</span>
                            <span class="status-tag ${statusClass}">${statusLabel}</span>
                        </div>
                        <div class="figure-meta-row">
                            <span>BBox: ${formula.bbox_width} × ${formula.bbox_height} pt</span>
                            <span style="font-weight:600; color:${ex ? '#34d399' : 'var(--primary)'};">${ex ? ex.filename : (formula.formula_type || '/Formula')}</span>
                        </div>
                        <div class="figure-alt-preview ${hasEffectiveAlt ? 'has-alt' : 'empty'}">
                            <div class="alt-label-bar">
                                <span>${isInjected ? '✓ INJECTED FORMULA ALT:' : (ex ? 'AUTHORITATIVE ALT TEXT FROM EXCEL:' : 'CURRENT /ALT OR /ACTUALTEXT:')}</span>
                                <div style="display:flex; gap:6px;">
                                    ${hasEffectiveAlt ? `<button class="btn-copy-alt-mini" onclick="event.stopPropagation(); window.copyFormulaAlt(${formula.formula_id})">Copy</button>` : ''}
                                    ${hasEffectiveAlt ? `<button class="btn-inject-mini" onclick="event.stopPropagation(); window.injectAltForFormula(${formula.formula_id})" title="Inject into PDF StructTree">⚡ Inject</button>` : ''}
                                </div>
                            </div>
                            <div class="alt-body-text">${escapeHtml(effectiveAlt || 'No /Alt or /ActualText defined in StructTree.')}</div>
                        </div>
                        <div class="card-footer-row">
                            <span class="footer-hint">${ex ? `Row ${ex.row} • Sr. ${ex.sr_no}` : `Page ${formula.page_number}`}</span>
                            <div style="display:flex; gap:6px;">
                                ${hasEffectiveAlt ? `
                                <button class="btn btn-inject btn-sm" onclick="event.stopPropagation(); window.injectAltForFormula(${formula.formula_id})" style="padding: 5px 10px; font-size: 0.78rem;">
                                    ⚡ Inject Alt
                                </button>` : ''}
                                <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); window.openFormulaModalById(${formula.formula_id})">
                                    Inspect Details
                                </button>
                            </div>
                        </div>
                    </div>
                `;
                formulasGrid.appendChild(card);
            });
        }
    }

    // Render Table
    if (formulasTableBody) {
        formulasTableBody.innerHTML = '';
        pageRecords.forEach(formula => {
            const isSelected = selectedFormulaIds.has(formula.formula_id);
            const tr = document.createElement('tr');
            const ex = formula.excel_match;
            const effectiveAlt = (ex && ex.alt_text) || formula.alt_text || formula.actual_text || '';
            const isMissing = !effectiveAlt;
            const isInjected = formula.status_label === 'Injected';
            const statusClass = isInjected ? 'present' : (ex ? 'present' : (formula.has_alt ? 'present' : 'missing'));
            const statusLabel = isInjected ? '✓ Alt Injected' : (ex ? `✓ Alt Linked (Row ${ex.row})` : (formula.has_alt ? 'Has /Alt' : 'Missing /Alt'));
            const mcidText = formula.mcids && formula.mcids.length > 0 ? formula.mcids.join(', ') : 'None';
            const altSummary = effectiveAlt || 'None';

            tr.innerHTML = `
                <td style="text-align: center;">
                    <input type="checkbox" class="table-card-checkbox" data-formula-id="${formula.formula_id}" ${isSelected ? 'checked' : ''} onchange="event.stopPropagation(); window.toggleFormulaSelection(${formula.formula_id}, this.checked)">
                </td>
                <td><strong>#${formula.formula_id}</strong></td>
                <td>
                    <div class="table-thumb">
                        <img src="${formula.image_url || ''}" alt="Formula ${formula.formula_id}" loading="lazy">
                    </div>
                </td>
                <td>Page ${formula.page_number || '?'}</td>
                <td><code>${mcidText}</code></td>
                <td>${formula.bbox_width} × ${formula.bbox_height}</td>
                <td><span class="status-tag ${statusClass}">${statusLabel}</span></td>
                <td><span style="font-size:0.8rem; color:${ex ? '#34d399' : 'var(--text-main)'};">${escapeHtml(altSummary.substring(0, 70))}...</span></td>
                <td>
                    <div style="display:flex; gap:6px;">
                        ${effectiveAlt ? `
                        <button class="btn btn-inject btn-sm" onclick="event.stopPropagation(); window.injectAltForFormula(${formula.formula_id})" style="padding: 4px 8px; font-size: 0.75rem;">
                            ⚡ Inject
                        </button>` : ''}
                        <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); window.openFormulaModalById(${formula.formula_id})">
                            Inspect
                        </button>
                    </div>
                </td>
            `;
            tr.addEventListener('click', () => openFormulaModal(formula));
            formulasTableBody.appendChild(tr);
        });
    }

    updateSelectionUI();
}

// ==========================================
// EXCEL MANIFEST RENDERING & PAGINATION
// ==========================================
function renderExcel() {
    const q = searchInput.value.toLowerCase().trim();

    filteredExcelRecords = currentExcelRecords.filter(rec => {
        if (activeFilter === 'missing' && rec.has_alt) return false;
        if (activeFilter === 'has_alt' && !rec.has_alt) return false;
        if (activeFilter === 'has_image' && !rec.has_image) return false;

        if (q) {
            const rowMatch = String(rec.row).includes(q);
            const srMatch = String(rec.sr_no || '').includes(q);
            const fnMatch = (rec.filename || '').toLowerCase().includes(q);
            const altMatch = (rec.alt_text || '').toLowerCase().includes(q);
            return rowMatch || srMatch || fnMatch || altMatch;
        }
        return true;
    });

    const totalFiltered = filteredExcelRecords.length;
    const totalPages = Math.ceil(totalFiltered / excelPageSize) || 1;
    if (excelPage > totalPages) excelPage = totalPages;
    if (excelPage < 1) excelPage = 1;

    const startIdx = (excelPage - 1) * excelPageSize;
    const endIdx = Math.min(startIdx + excelPageSize, totalFiltered);
    const pageRecords = filteredExcelRecords.slice(startIdx, endIdx);

    // Update Pagination UI
    paginationInfo.textContent = totalFiltered > 0
        ? `Showing ${(startIdx + 1).toLocaleString()} - ${endIdx.toLocaleString()} of ${totalFiltered.toLocaleString()} Excel items`
        : 'No records found';
    pageCurrentDisplay.textContent = `Page ${excelPage.toLocaleString()} of ${totalPages.toLocaleString()}`;

    pageFirstBtn.disabled = excelPage <= 1;
    pagePrevBtn.disabled = excelPage <= 1;
    pageNextBtn.disabled = excelPage >= totalPages;
    pageLastBtn.disabled = excelPage >= totalPages;

    // Render Excel Grid
    excelGrid.innerHTML = '';
    if (pageRecords.length === 0) {
        excelGrid.innerHTML = `
            <div style="grid-column: 1 / -1; text-align: center; padding: 60px 20px; color: var(--text-dim);">
                <h3>No Excel manifest records match the filter or search</h3>
            </div>
        `;
    } else {
        pageRecords.forEach(rec => {
            const card = document.createElement('div');
            card.className = 'excel-card';
            card.addEventListener('click', () => openExcelModal(rec));

            const hasImg = rec.has_image && rec.image_url;
            const altLower = (rec.alt_text || '').toLowerCase();
            const isSymbol = altLower.includes('checkbox') || altLower.includes('flowchart') || altLower.includes('arrow') || altLower.includes('circle');
            
            let imgHtml = '';
            if (hasImg) {
                imgHtml = `<img src="${rec.image_url}" alt="Row ${rec.row}" loading="lazy" onerror="this.onerror=null; this.style.display='none'; this.nextElementSibling.style.display='flex';"><div class="excel-no-image-placeholder" style="display:none;"><svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><line x1="21" y1="21" x2="3" y2="3"/></svg><span>Text-Only Alt Record</span></div>`;
            } else if (isSymbol) {
                imgHtml = `<div class="excel-no-image-placeholder symbol-placeholder">
                             <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="#34d399" stroke-width="1.8"><rect x="3" y="3" width="18" height="18" rx="3"/><path d="m9 12 2 2 4-4"/></svg>
                             <span style="color:#34d399; font-weight:600; font-size:0.8rem;">Form / Layout Symbol</span>
                             <span style="font-size:0.75rem; color:var(--text-dim); text-align:center; padding:0 8px;">${escapeHtml(rec.alt_text)}</span>
                           </div>`;
            } else {
                imgHtml = `<div class="excel-no-image-placeholder">
                             <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><line x1="21" y1="21" x2="3" y2="3"/></svg>
                             <span>Text-Only Alt Record</span>
                           </div>`;
            }

            const statusClass = rec.has_alt ? 'authoritative' : 'missing';
            const statusLabel = rec.has_alt ? 'Authoritative Alt' : 'Missing Alt';

            card.innerHTML = `
                <div class="excel-image-wrapper">
                    <span class="row-chip">Row ${rec.row} ${rec.sr_no && String(rec.sr_no).length <= 12 ? '(Sr. ' + escapeHtml(String(rec.sr_no)) + ')' : ''}</span>
                    ${imgHtml}
                </div>
                <div class="excel-card-content">
                    <div class="excel-title-row">
                        <span class="excel-filename" title="${escapeHtml(rec.filename || '')}">${escapeHtml(rec.filename || `Row_${rec.row}`)}</span>
                        <span class="status-tag ${statusClass}">${statusLabel}</span>
                    </div>
                    <div class="excel-alt-box-card ${rec.alt_text ? '' : 'empty'}" title="${escapeHtml(rec.alt_text || '')}">
                        ${escapeHtml(rec.alt_text || 'No ALT text provided in Excel manifest.')}
                    </div>
                    <div class="excel-card-actions">
                        <button class="btn-copy-alt" onclick="event.stopPropagation(); window.copyExcelAlt(${rec.row})">
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                            Copy Alt
                        </button>
                        <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); window.openExcelModalByRow(${rec.row})">
                            Inspect
                        </button>
                    </div>
                </div>
            `;
            excelGrid.appendChild(card);
        });
    }

    // Render Excel Table
    excelTableBody.innerHTML = '';
    pageRecords.forEach(rec => {
        const tr = document.createElement('tr');
        const statusClass = rec.has_alt ? 'authoritative' : 'missing';
        const statusLabel = rec.has_alt ? 'Authoritative' : 'Missing';
        const hasImg = rec.has_image && rec.image_url;

        tr.innerHTML = `
            <td><strong>Row ${rec.row}</strong></td>
            <td>
                <div class="table-thumb">
                    ${hasImg ? `<img src="${rec.image_url}" alt="Row ${rec.row}" loading="lazy" onerror="this.parentElement.innerHTML='<span style=\\'font-size:0.7rem;color:var(--text-dim);\\'>No Img</span>';">` : '<span style="font-size:0.7rem;color:var(--text-dim);">No Img</span>'}
                </div>
            </td>
            <td><code>${escapeHtml(rec.filename || '')}</code></td>
            <td>Sr. ${rec.sr_no}</td>
            <td><span class="status-tag ${statusClass}">${statusLabel}</span></td>
            <td><span style="font-size:0.8rem; color:var(--text-main);">${escapeHtml((rec.alt_text || '').substring(0, 80))}...</span></td>
            <td>
                <div style="display:flex; gap:6px;">
                    <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); window.openExcelModalByRow(${rec.row})">Inspect</button>
                    <button class="btn-copy-alt" onclick="event.stopPropagation(); window.copyExcelAlt(${rec.row})">Copy</button>
                </div>
            </td>
        `;
        tr.addEventListener('click', () => openExcelModal(rec));
        excelTableBody.appendChild(tr);
    });
}

// Quick helper for copying Alt by row / figure / formula
window.copyExcelAlt = function(row) {
    const rec = currentExcelRecords.find(r => r.row === row);
    if (rec && rec.alt_text) {
        copyToClipboard(rec.alt_text, `Copied Row ${row} ALT text!`);
    } else {
        showToast(`Row ${row} has no ALT text.`);
    }
};

window.copyFigureAlt = function(figId) {
    const fig = currentFigures.find(f => f.figure_id === figId);
    if (!fig) return;
    const text = (fig.excel_match && fig.excel_match.alt_text) || fig.alt_text;
    if (text) {
        copyToClipboard(text, `Copied Alt for Figure ${figId}!`);
    } else {
        showToast(`Figure ${figId} has no ALT text.`);
    }
};

window.copyFormulaAlt = function(formulaId) {
    const formula = currentFormulas.find(f => f.formula_id === formulaId);
    if (!formula) return;
    const text = (formula.excel_match && formula.excel_match.alt_text) || formula.alt_text || formula.actual_text;
    if (text) {
        copyToClipboard(text, `Copied Alt for Formula #${formulaId}!`);
    } else {
        showToast(`Formula #${formulaId} has no ALT text.`);
    }
};

window.openExcelModalByRow = function(row) {
    const rec = currentExcelRecords.find(r => r.row === row);
    if (rec) openExcelModal(rec);
};

// ==========================================
// ALT TEXT INJECTION FUNCTIONS
// ==========================================
window.injectAltForFigure = async function(figId, customAlt) {
    if (!currentSession || !currentSession.session_id) {
        showToast('Please load or upload a PDF first.');
        return;
    }

    const fig = currentFigures.find(f => f.figure_id === figId);
    if (!fig) return;

    let altToInject = (customAlt !== undefined && customAlt !== null) ? customAlt : ((fig.excel_match && fig.excel_match.alt_text) || fig.alt_text || '');
    altToInject = String(altToInject).trim();

    if (!altToInject) {
        showToast(`No ALT text available to inject for Figure ${figId}.`);
        return;
    }

    showToast(`Injecting ALT into StructTree for Figure ${figId}...`);

    try {
        const res = await fetch(`/api/inject-single-alt/${currentSession.session_id}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                figure_id: figId,
                alt_text: altToInject
            })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Injection failed');
        }

        const data = await res.json();
        fig.alt_text = data.injected_alt;
        fig.has_alt = true;
        fig.status_label = 'Injected';

        if (currentSession) {
            currentSession.has_injected_pdf = true;
            currentSession.has_alt_count = data.has_alt_count;
            currentSession.missing_alt_count = data.missing_alt_count;
        }

        statHasAlt.textContent = data.has_alt_count;
        statMissingAlt.textContent = data.missing_alt_count;

        if (downloadInjectedPdfBtn) {
            downloadInjectedPdfBtn.href = data.download_url;
            downloadInjectedPdfBtn.style.display = 'inline-flex';
        }

        if (figureModal && figureModal.style.display === 'flex' && currentModalFigure && currentModalFigure.figure_id === figId) {
            if (modalAltTextarea) modalAltTextarea.value = data.injected_alt;
            if (modalInjectBtn) modalInjectBtn.textContent = '✓ Alt Injected';
        }

        showToast(`✓ Injected ALT into Figure ${figId}! PDF ready for download.`);
        refreshActiveView();

    } catch (err) {
        alert('Injection error: ' + err.message);
    }
};

window.injectAltForFormula = async function(formulaId, customAlt) {
    if (!currentSession || !currentSession.session_id) {
        showToast('Please load or upload a PDF first.');
        return;
    }

    const formula = currentFormulas.find(f => f.formula_id === formulaId);
    if (!formula) return;

    let altToInject = (customAlt !== undefined && customAlt !== null) ? customAlt : ((formula.excel_match && formula.excel_match.alt_text) || formula.alt_text || formula.actual_text || '');
    altToInject = String(altToInject).trim();

    if (!altToInject) {
        showToast(`No ALT text available to inject for Formula #${formulaId}.`);
        return;
    }

    showToast(`Injecting ALT into StructTree for Formula #${formulaId}...`);

    try {
        const res = await fetch(`/api/inject-single-formula-alt/${currentSession.session_id}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                formula_id: formulaId,
                alt_text: altToInject
            })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Formula injection failed');
        }

        const data = await res.json();
        formula.alt_text = data.injected_alt;
        formula.has_alt = true;
        formula.status_label = 'Injected';

        if (currentSession) {
            currentSession.has_injected_pdf = true;
            currentSession.has_formula_alt_count = data.has_formula_alt_count;
            currentSession.missing_formula_alt_count = data.missing_formula_alt_count;
        }

        if (statFormulaHasAlt) statFormulaHasAlt.textContent = data.has_formula_alt_count;
        if (statFormulaMissingAlt) statFormulaMissingAlt.textContent = data.missing_formula_alt_count;

        if (downloadInjectedPdfBtn) {
            downloadInjectedPdfBtn.href = data.download_url;
            downloadInjectedPdfBtn.style.display = 'inline-flex';
        }

        if (figureModal && figureModal.style.display === 'flex' && currentModalFormula && currentModalFormula.formula_id === formulaId) {
            if (modalAltTextarea) modalAltTextarea.value = data.injected_alt;
            if (modalInjectBtn) modalInjectBtn.textContent = '✓ Alt Injected';
        }

        showToast(`✓ Injected ALT into Formula #${formulaId}! PDF ready for download.`);
        refreshActiveView();

    } catch (err) {
        alert('Formula injection error: ' + err.message);
    }
};

async function handleBatchInjectAlt() {
    if (!currentSession || !currentSession.session_id) {
        showToast('Please load or upload a PDF first.');
        return;
    }

    if (currentSourceTab === 'formula') {
        if (!currentFormulas || currentFormulas.length === 0) {
            showToast('No formulas found in current PDF.');
            return;
        }

        const canInject = currentFormulas.filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text || f.actual_text);
        if (canInject.length === 0) {
            showToast('No authoritative ALT texts found to inject for formulas.');
            return;
        }

        const origBtnHtml = injectAltBtn.innerHTML;
        injectAltBtn.innerHTML = `
            <div class="spinner" style="width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; display:inline-block; vertical-align:middle; margin-right:6px;"></div>
            <span>Injecting Formula Alt into PDF...</span>
        `;
        injectAltBtn.disabled = true;

        try {
            const res = await fetch(`/api/inject-formula-alt/${currentSession.session_id}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({})
            });

            if (!res.ok) {
                const err = await res.json();
                throw new Error(err.detail || 'Batch formula injection failed');
            }

            const data = await res.json();
            currentFormulas = data.formulas;
            if (currentSession) {
                currentSession.formulas = data.formulas;
                currentSession.has_formula_alt_count = data.has_formula_alt_count;
                currentSession.missing_formula_alt_count = data.missing_formula_alt_count;
                currentSession.has_injected_pdf = true;
            }

            if (statFormulaHasAlt) statFormulaHasAlt.textContent = data.has_formula_alt_count;
            if (statFormulaMissingAlt) statFormulaMissingAlt.textContent = data.missing_formula_alt_count;

            if (downloadInjectedPdfBtn) {
                downloadInjectedPdfBtn.href = data.download_url;
                downloadInjectedPdfBtn.style.display = 'inline-flex';
            }

            injectAltBtn.innerHTML = `
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>
                <span>✓ Injected All (${data.injected_count})</span>
            `;
            injectAltBtn.disabled = false;

            showToast(`⚡ Successfully injected ${data.injected_count} ALT texts into StructTreeRoot /Formula tags!`);
            refreshActiveView();

        } catch (err) {
            injectAltBtn.innerHTML = origBtnHtml;
            injectAltBtn.disabled = false;
            alert('Formula injection failed: ' + err.message);
        }
        return;
    }

    if (!currentFigures || currentFigures.length === 0) {
        showToast('No figures found in current PDF.');
        return;
    }

    const canInject = currentFigures.filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text);
    if (canInject.length === 0) {
        showToast('No authoritative ALT texts found to inject.');
        return;
    }

    const origBtnHtml = injectAltBtn.innerHTML;
    injectAltBtn.innerHTML = `
        <div class="spinner" style="width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; display:inline-block; vertical-align:middle; margin-right:6px;"></div>
        <span>Injecting Alt into PDF...</span>
    `;
    injectAltBtn.disabled = true;

    try {
        const res = await fetch(`/api/inject-alt/${currentSession.session_id}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({})
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Batch injection failed');
        }

        const data = await res.json();
        currentFigures = data.figures;
        if (currentSession) {
            currentSession.figures = data.figures;
            currentSession.has_alt_count = data.has_alt_count;
            currentSession.missing_alt_count = data.missing_alt_count;
            currentSession.has_injected_pdf = true;
        }

        statHasAlt.textContent = data.has_alt_count;
        statMissingAlt.textContent = data.missing_alt_count;

        if (downloadInjectedPdfBtn) {
            downloadInjectedPdfBtn.href = data.download_url;
            downloadInjectedPdfBtn.style.display = 'inline-flex';
        }

        injectAltBtn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>
            <span>✓ Injected All (${data.injected_count})</span>
        `;
        injectAltBtn.disabled = false;

        showToast(`⚡ Successfully injected ${data.injected_count} ALT texts into StructTreeRoot /Figure tags!`);
        refreshActiveView();

    } catch (err) {
        injectAltBtn.innerHTML = origBtnHtml;
        injectAltBtn.disabled = false;
        alert('Injection failed: ' + err.message);
    }
}

async function downloadUnselectedFigures() {
    if (!currentSession || !currentSession.session_id) {
        showToast('Please load or upload a PDF first.');
        return;
    }

    if (!currentFigures || currentFigures.length === 0) {
        showToast('No figures found in current PDF.');
        return;
    }

    const unselectedFigureIds = currentFigures
        .filter(fig => !selectedFigureIds.has(fig.figure_id))
        .map(fig => fig.figure_id);

    if (unselectedFigureIds.length === 0) {
        showToast('All figures are selected. Please deselect some figures to download.');
        return;
    }

    try {
        const res = await fetch(`/api/download-unselected-excel/${currentSession.session_id}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ figure_ids: unselectedFigureIds })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Download failed');
        }

        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `unselected_figures_${currentSession.session_id.slice(0, 8)}.xlsx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);

        showToast(`✓ Downloaded ${unselectedFigureIds.length} unselected figures as Excel!`);
    } catch (err) {
        alert('Download failed: ' + err.message);
    }
}

// ==========================================
// SELECTION STATE MANAGEMENT & INJECTION
// ==========================================
function updateSelectionUI() {
    if (currentSourceTab === 'formula') {
        const selectableFormulas = currentFormulas.filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text || f.actual_text);
        const totalSelectable = selectableFormulas.length;
        const selectedCount = selectedFormulaIds.size;

        if (selectedCountBadge) selectedCountBadge.textContent = selectedCount;
        if (totalSelectableBadge) totalSelectableBadge.textContent = totalSelectable;

        if (injectSelectedBtn) {
            injectSelectedBtn.disabled = selectedCount === 0;
        }
        if (injectSelectedBtnText) {
            injectSelectedBtnText.textContent = `⚡ Inject Selected Alt into PDF (${selectedCount})`;
        }
        if (removeAltBtnText) {
            if (selectedCount > 0) {
                removeAltBtnText.textContent = `🗑️ Remove Selected Formula Alt (${selectedCount})`;
            } else {
                removeAltBtnText.textContent = `🗑️ Remove All Formula Alt`;
            }
        }

        const allChecked = totalSelectable > 0 && selectedCount === totalSelectable;
        const isIndeterminate = selectedCount > 0 && selectedCount < totalSelectable;

        if (selectAllCheckbox) {
            selectAllCheckbox.checked = allChecked;
            selectAllCheckbox.indeterminate = isIndeterminate;
        }
        if (formulaTableSelectAllCheckbox) {
            formulaTableSelectAllCheckbox.checked = allChecked;
            formulaTableSelectAllCheckbox.indeterminate = isIndeterminate;
        }

        document.querySelectorAll('.figure-card[data-formula-id]').forEach(card => {
            const fid = parseInt(card.getAttribute('data-formula-id'), 10);
            const isSel = selectedFormulaIds.has(fid);
            card.classList.toggle('selected', isSel);

            const cb = card.querySelector('.card-checkbox');
            if (cb) cb.checked = isSel;

            const btn = card.querySelector('.btn-card-select-toggle');
            if (btn) {
                btn.classList.toggle('selected', isSel);
                btn.textContent = isSel ? '✓ Selected' : '+ Select';
            }
        });

        document.querySelectorAll('.table-card-checkbox[data-formula-id]').forEach(cb => {
            const fid = parseInt(cb.getAttribute('data-formula-id'), 10);
            cb.checked = selectedFormulaIds.has(fid);
        });
        return;
    }

    const selectableFigures = currentFigures.filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text);
    const totalSelectable = selectableFigures.length;
    const selectedCount = selectedFigureIds.size;

    if (selectedCountBadge) selectedCountBadge.textContent = selectedCount;
    if (totalSelectableBadge) totalSelectableBadge.textContent = totalSelectable;

    if (injectSelectedBtn) {
        injectSelectedBtn.disabled = selectedCount === 0;
    }
    if (injectSelectedBtnText) {
        injectSelectedBtnText.textContent = `⚡ Inject Selected Alt into PDF (${selectedCount})`;
    }
    if (removeAltBtnText) {
        if (selectedCount > 0) {
            removeAltBtnText.textContent = `🗑️ Remove Selected Alt (${selectedCount})`;
        } else {
            removeAltBtnText.textContent = `🗑️ Remove All Alt Text`;
        }
    }

    // Sync select all checkboxes
    const allChecked = totalSelectable > 0 && selectedCount === totalSelectable;
    const isIndeterminate = selectedCount > 0 && selectedCount < totalSelectable;

    if (selectAllCheckbox) {
        selectAllCheckbox.checked = allChecked;
        selectAllCheckbox.indeterminate = isIndeterminate;
    }
    if (tableSelectAllCheckbox) {
        tableSelectAllCheckbox.checked = allChecked;
        tableSelectAllCheckbox.indeterminate = isIndeterminate;
    }

    // Sync each figure card in the DOM
    document.querySelectorAll('.figure-card[data-fig-id]').forEach(card => {
        const figId = parseInt(card.getAttribute('data-fig-id'), 10);
        const isSel = selectedFigureIds.has(figId);
        card.classList.toggle('selected', isSel);

        const cb = card.querySelector('.card-checkbox');
        if (cb) cb.checked = isSel;

        const btn = card.querySelector('.btn-card-select-toggle');
        if (btn) {
            btn.classList.toggle('selected', isSel);
            btn.textContent = isSel ? '✓ Selected' : '+ Select';
        }
    });

    // Sync match cards in the DOM
    document.querySelectorAll('.match-card[data-fig-id]').forEach(card => {
        const figId = parseInt(card.getAttribute('data-fig-id'), 10);
        const isSel = selectedFigureIds.has(figId);
        card.classList.toggle('selected', isSel);

        const cb = card.querySelector('.match-card-checkbox');
        if (cb) cb.checked = isSel;

        const btn = card.querySelector('.btn-card-select-toggle');
        if (btn) {
            btn.classList.toggle('selected', isSel);
            btn.textContent = isSel ? '✓ Selected' : '+ Select';
        }
    });

    // Sync table rows
    document.querySelectorAll('.table-card-checkbox[data-fig-id]').forEach(cb => {
        const figId = parseInt(cb.getAttribute('data-fig-id'), 10);
        cb.checked = selectedFigureIds.has(figId);
    });
}

window.toggleFigureSelection = function(figId, forceState) {
    const id = parseInt(figId, 10);
    if (forceState !== undefined) {
        if (forceState) selectedFigureIds.add(id);
        else selectedFigureIds.delete(id);
    } else {
        if (selectedFigureIds.has(id)) selectedFigureIds.delete(id);
        else selectedFigureIds.add(id);
    }
    updateSelectionUI();
};

window.selectAllFigures = function() {
    currentFigures.forEach(f => {
        const hasAlt = (f.excel_match && f.excel_match.alt_text) || f.alt_text;
        if (hasAlt) {
            selectedFigureIds.add(f.figure_id);
        }
    });
    updateSelectionUI();
    showToast(`Selected all ${selectedFigureIds.size} figure boxes for injection.`);
};

window.deselectAllFigures = function() {
    selectedFigureIds.clear();
    updateSelectionUI();
    showToast('Deselected all figures.');
};

window.toggleFormulaSelection = function(formulaId, forceState) {
    const id = parseInt(formulaId, 10);
    if (forceState !== undefined) {
        if (forceState) selectedFormulaIds.add(id);
        else selectedFormulaIds.delete(id);
    } else {
        if (selectedFormulaIds.has(id)) selectedFormulaIds.delete(id);
        else selectedFormulaIds.add(id);
    }
    updateSelectionUI();
};

window.selectAllFormulas = function() {
    currentFormulas.forEach(f => {
        const hasAlt = (f.excel_match && f.excel_match.alt_text) || f.alt_text || f.actual_text;
        if (hasAlt) {
            selectedFormulaIds.add(f.formula_id);
        }
    });
    updateSelectionUI();
    showToast(`Selected all ${selectedFormulaIds.size} formula boxes for injection.`);
};

window.deselectAllFormulas = function() {
    selectedFormulaIds.clear();
    updateSelectionUI();
    showToast('Deselected all formulas.');
};

async function handleInjectSelected() {
    if (!currentSession || !currentSession.session_id) {
        showToast('Please load or upload a PDF first.');
        return;
    }

    if (currentSourceTab === 'formula') {
        if (selectedFormulaIds.size === 0) {
            showToast('Please select at least one formula box to inject.');
            return;
        }

        const injections = {};
        selectedFormulaIds.forEach(id => {
            const formula = currentFormulas.find(f => f.formula_id === id);
            if (formula) {
                const alt = (formula.excel_match && formula.excel_match.alt_text) || formula.alt_text || formula.actual_text;
                if (alt) injections[id] = alt.trim();
            }
        });

        if (Object.keys(injections).length === 0) {
            showToast('None of the selected formulas have ALT text available.');
            return;
        }

        const origBtnHtml = injectSelectedBtn.innerHTML;
        injectSelectedBtn.innerHTML = `
            <div class="spinner" style="width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; display:inline-block; vertical-align:middle; margin-right:6px;"></div>
            <span>Injecting (${selectedFormulaIds.size} Formulas)...</span>
        `;
        injectSelectedBtn.disabled = true;

        try {
            const res = await fetch(`/api/inject-formula-alt/${currentSession.session_id}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ injections: injections })
            });

            if (!res.ok) {
                const err = await res.json();
                throw new Error(err.detail || 'Formula injection failed');
            }

            const data = await res.json();
            currentFormulas = data.formulas;
            if (currentSession) {
                currentSession.formulas = data.formulas;
                currentSession.has_formula_alt_count = data.has_formula_alt_count;
                currentSession.missing_formula_alt_count = data.missing_formula_alt_count;
                currentSession.has_injected_pdf = true;
            }

            if (statFormulaHasAlt) statFormulaHasAlt.textContent = data.has_formula_alt_count;
            if (statFormulaMissingAlt) statFormulaMissingAlt.textContent = data.missing_formula_alt_count;

            if (downloadInjectedPdfBtn) {
                downloadInjectedPdfBtn.href = data.download_url;
                downloadInjectedPdfBtn.style.display = 'inline-flex';
            }

            showToast(`⚡ Successfully injected ${data.injected_count} selected formula ALT texts into StructTreeRoot!`);
            refreshActiveView();
            updateSelectionUI();

        } catch (err) {
            injectSelectedBtn.innerHTML = origBtnHtml;
            injectSelectedBtn.disabled = false;
            alert('Formula injection failed: ' + err.message);
        }
        return;
    }

    if (selectedFigureIds.size === 0) {
        showToast('Please select at least one figure box to inject.');
        return;
    }

    const injections = {};
    selectedFigureIds.forEach(id => {
        const fig = currentFigures.find(f => f.figure_id === id);
        if (fig) {
            const alt = (fig.excel_match && fig.excel_match.alt_text) || fig.alt_text;
            if (alt) injections[id] = alt.trim();
        }
    });

    if (Object.keys(injections).length === 0) {
        showToast('None of the selected figures have ALT text available.');
        return;
    }

    const origBtnHtml = injectSelectedBtn.innerHTML;
    injectSelectedBtn.innerHTML = `
        <div class="spinner" style="width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; display:inline-block; vertical-align:middle; margin-right:6px;"></div>
        <span>Injecting (${selectedFigureIds.size} Selected)...</span>
    `;
    injectSelectedBtn.disabled = true;

    try {
        const res = await fetch(`/api/inject-alt/${currentSession.session_id}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ injections: injections })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Injection failed');
        }

        const data = await res.json();
        currentFigures = data.figures;
        if (currentSession) {
            currentSession.figures = data.figures;
            currentSession.has_alt_count = data.has_alt_count;
            currentSession.missing_alt_count = data.missing_alt_count;
            currentSession.has_injected_pdf = true;
        }

        statHasAlt.textContent = data.has_alt_count;
        statMissingAlt.textContent = data.missing_alt_count;

        if (downloadInjectedPdfBtn) {
            downloadInjectedPdfBtn.href = data.download_url;
            downloadInjectedPdfBtn.style.display = 'inline-flex';
        }

        showToast(`⚡ Successfully injected ${data.injected_count} selected ALT texts into StructTreeRoot!`);
        refreshActiveView();
        updateSelectionUI();

    } catch (err) {
        injectSelectedBtn.innerHTML = origBtnHtml;
        injectSelectedBtn.disabled = false;
        alert('Injection failed: ' + err.message);
    }
}

async function handleRemoveAlt() {
    if (!currentSession || !currentSession.session_id) {
        showToast('Please load or upload a PDF first.');
        return;
    }

    if (currentSourceTab === 'formula') {
        if (!currentFormulas || currentFormulas.length === 0) {
            showToast('No formulas found in current PDF.');
            return;
        }

        const selectedCount = selectedFormulaIds.size;
        const isTargeted = selectedCount > 0;
        const targetIds = isTargeted ? Array.from(selectedFormulaIds) : null;

        const confirmMsg = isTargeted
            ? `Are you sure you want to remove /Alt accessibility text from the ${selectedCount} selected formula(s) in the PDF?`
            : `Are you sure you want to remove /Alt accessibility text from ALL formulas in the PDF?`;

        if (!confirm(confirmMsg)) {
            return;
        }

        const origBtnHtml = removeAltBtn.innerHTML;
        removeAltBtn.innerHTML = `
            <div class="spinner" style="width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; display:inline-block; vertical-align:middle; margin-right:6px;"></div>
            <span>Removing Formula Alt...</span>
        `;
        removeAltBtn.disabled = true;

        try {
            const res = await fetch(`/api/remove-formula-alt/${currentSession.session_id}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ formula_ids: targetIds })
            });

            if (!res.ok) {
                const err = await res.json();
                throw new Error(err.detail || 'Formula removal failed');
            }

            const data = await res.json();
            currentFormulas = data.formulas;
            if (currentSession) {
                currentSession.formulas = data.formulas;
                currentSession.has_formula_alt_count = data.has_formula_alt_count;
                currentSession.missing_formula_alt_count = data.missing_formula_alt_count;
                currentSession.has_injected_pdf = true;
            }

            if (statFormulaHasAlt) statFormulaHasAlt.textContent = data.has_formula_alt_count;
            if (statFormulaMissingAlt) statFormulaMissingAlt.textContent = data.missing_formula_alt_count;

            if (downloadInjectedPdfBtn) {
                downloadInjectedPdfBtn.href = data.download_url;
                downloadInjectedPdfBtn.style.display = 'inline-flex';
            }

            showToast(`🗑️ Successfully removed /Alt text from ${data.removed_count} formula(s) in the PDF!`);
            refreshActiveView();
            updateSelectionUI();

        } catch (err) {
            alert('Formula removal failed: ' + err.message);
        } finally {
            removeAltBtn.innerHTML = origBtnHtml;
            removeAltBtn.disabled = false;
        }
        return;
    }

    if (!currentFigures || currentFigures.length === 0) {
        showToast('No figures found in current PDF.');
        return;
    }

    const selectedCount = selectedFigureIds.size;
    const isTargeted = selectedCount > 0;
    const targetIds = isTargeted ? Array.from(selectedFigureIds) : null;

    const confirmMsg = isTargeted
        ? `Are you sure you want to remove /Alt accessibility text from the ${selectedCount} selected figure(s) in the PDF?`
        : `Are you sure you want to remove /Alt accessibility text from ALL figures in the PDF?`;

    if (!confirm(confirmMsg)) {
        return;
    }

    const origBtnHtml = removeAltBtn.innerHTML;
    removeAltBtn.innerHTML = `
        <div class="spinner" style="width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; display:inline-block; vertical-align:middle; margin-right:6px;"></div>
        <span>Removing Alt...</span>
    `;
    removeAltBtn.disabled = true;

    try {
        const res = await fetch(`/api/remove-alt/${currentSession.session_id}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ figure_ids: targetIds })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Removal failed');
        }

        const data = await res.json();
        currentFigures = data.figures;
        if (currentSession) {
            currentSession.figures = data.figures;
            currentSession.has_alt_count = data.has_alt_count;
            currentSession.missing_alt_count = data.missing_alt_count;
            currentSession.has_injected_pdf = true;
        }

        statHasAlt.textContent = data.has_alt_count;
        statMissingAlt.textContent = data.missing_alt_count;

        if (downloadInjectedPdfBtn) {
            downloadInjectedPdfBtn.href = data.download_url;
            downloadInjectedPdfBtn.style.display = 'inline-flex';
        }

        showToast(`🗑️ Successfully removed /Alt text from ${data.removed_count} figure(s) in the PDF!`);
        refreshActiveView();
        updateSelectionUI();

    } catch (err) {
        alert('Removal failed: ' + err.message);
    } finally {
        removeAltBtn.innerHTML = origBtnHtml;
        removeAltBtn.disabled = false;
    }
}

// ==========================================
// MATCHED SIDE-BY-SIDE VIEW RENDERING
// ==========================================
function renderMatched() {
    matchGrid.innerHTML = '';
    const q = searchInput.value.toLowerCase().trim();

    const matched = currentFigures.filter(f => f.excel_match).filter(fig => {
        if (!q) return true;
        const pageMatch = String(fig.page_number).includes(q);
        const idMatch = String(fig.figure_id).includes(q);
        const altMatch = (fig.excel_match.alt_text || '').toLowerCase().includes(q);
        return pageMatch || idMatch || altMatch;
    });

    if (matched.length === 0) {
        matchGrid.innerHTML = `
            <div style="text-align:center; padding:60px 20px; color:var(--text-dim);">
                <h3>No aligned figure matches found. Ensure both PDF and Excel manifest are loaded.</h3>
            </div>
        `;
        return;
    }

    matched.forEach(fig => {
        const ex = fig.excel_match;
        const confVal = (typeof fig.confidence === 'number' && !isNaN(fig.confidence)) ? fig.confidence : 0.95;
        const confPct = Math.round(confVal * 100);
        const isSelected = selectedFigureIds.has(fig.figure_id);

        const card = document.createElement('div');
        card.className = `match-card ${isSelected ? 'selected' : ''}`;
        card.setAttribute('data-fig-id', fig.figure_id);
        card.innerHTML = `
            <!-- PDF Figure Side -->
            <div class="match-side">
                <div class="match-side-header">
                    <label class="match-select-label" onclick="event.stopPropagation();" title="Select PDF Figure ${fig.figure_id} for injection">
                        <input type="checkbox" class="match-card-checkbox custom-checkbox" data-fig-id="${fig.figure_id}" ${isSelected ? 'checked' : ''} onchange="window.toggleFigureSelection(${fig.figure_id}, this.checked)">
                        <span class="match-side-title">PDF Figure ${fig.figure_id}</span>
                    </label>
                    <div style="display:flex; align-items:center; gap:8px;">
                        <button class="btn-card-select-toggle ${isSelected ? 'selected' : ''}" onclick="event.stopPropagation(); window.toggleFigureSelection(${fig.figure_id})" title="Toggle selection for injection">
                            ${isSelected ? '✓ Selected' : '+ Select'}
                        </button>
                        <span class="page-chip" style="position:static;">Page ${fig.page_number}</span>
                    </div>
                </div>
                <div class="match-img-frame">
                    <img src="${fig.image_url || ''}" alt="PDF Fig ${fig.figure_id}" loading="lazy">
                </div>
                <div class="figure-alt-preview ${fig.alt_text ? '' : 'empty'}">
                    <strong>Current PDF Alt:</strong> ${escapeHtml(fig.alt_text || 'None (Missing /Alt in PDF StructTree)')}
                </div>
            </div>

            <!-- Match Divider / Inject Action -->
            <div class="match-divider">
                <button class="btn btn-inject btn-sm" onclick="window.injectAltForFigure(${fig.figure_id})">
                    ⚡ Inject into PDF
                </button>
            </div>

            <!-- Excel Image Side -->
            <div class="match-side">
                <div class="match-side-header">
                    <span class="match-side-title">Excel Row ${ex.row} (Sr. ${ex.sr_no || ex.row - 1})</span>
                    <span class="row-chip" style="position:static;">${escapeHtml(ex.filename || '')}</span>
                </div>
                <div class="match-img-frame">
                    <img src="${ex.image_url || ''}" alt="Excel match" loading="lazy">
                </div>
                <div class="excel-alt-box-card" style="min-height:54px;">
                    <strong>Authoritative Alt:</strong> ${escapeHtml(ex.alt_text || 'No ALT text')}
                </div>
            </div>
        `;
        matchGrid.appendChild(card);
    });
}

// ==========================================
// MODAL DIALOGS
// ==========================================
window.openPdfModalById = function(figId) {
    const fig = currentFigures.find(f => f.figure_id === figId);
    if (fig) openPdfModal(fig);
};

function openPdfModal(fig) {
    currentModalFigure = fig;
    modalTitle.textContent = `PDF Figure ${fig.figure_id} Accessibility Inspector`;
    modalSubtitle.textContent = `Page ${fig.page_number} • Marked Content ID: ${fig.mcids && fig.mcids.length ? fig.mcids.join(', ') : 'None'}`;
    modalImage.src = fig.image_url || '';
    modalDownloadLink.href = fig.image_url || '';
    modalDownloadLink.setAttribute('download', fig.crop_filename || `figure_${fig.figure_id}.png`);

    modalTypeLabel.textContent = 'Accessibility Tag';
    modalTypeDisplay.textContent = '/S /Figure (Tagged Structural Element)';

    const ex = fig.excel_match;
    const effectiveAlt = (ex && ex.alt_text) || fig.alt_text || '';
    const isInjected = fig.status_label === 'Injected';

    modalAltLabel.textContent = 'Authoritative /Alt Attribute (Editable for PDF Injection)';
    if (modalAltTextarea) {
        modalAltTextarea.value = effectiveAlt;
        modalAltTextarea.readOnly = false;
        modalAltTextarea.style.display = 'block';
    }
    if (modalAltDisplay) {
        modalAltDisplay.style.display = 'none';
    }

    if (modalInjectBtn) {
        modalInjectBtn.style.display = 'inline-flex';
        modalInjectBtn.textContent = isInjected ? '⚡ Re-Inject into PDF' : '⚡ Inject into PDF';
    }

    modalBBoxGroup.style.display = 'block';
    modalBBoxDisplay.textContent = fig.bbox ? `[${fig.bbox.join(', ')}] (${fig.bbox_width} pt × ${fig.bbox_height} pt)` : 'No bounding box';
    modalPathGroup.style.display = 'block';
    modalPathDisplay.textContent = fig.path || 'Root';
    modalMcidGroup.style.display = 'block';
    modalMcidDisplay.textContent = fig.mcids && fig.mcids.length ? `[${fig.mcids.join(', ')}]` : 'None';

    // Matched Excel preview
    if (ex) {
        modalMatchedExcelBox.style.display = 'block';
        if (modalMatchScore) modalMatchScore.style.display = 'none';
        modalExcelImage.src = ex.image_url || '';
        modalExcelRowMeta.textContent = `Row ${ex.row} (Sr. No. ${ex.sr_no || ex.row - 1}) • ${ex.filename || ''}`;
        modalExcelAltText.textContent = ex.alt_text || 'No ALT text available.';
    } else {
        modalMatchedExcelBox.style.display = 'none';
    }

    figureModal.style.display = 'flex';
}

window.openFormulaModalById = function(formulaId) {
    const formula = currentFormulas.find(f => f.formula_id === formulaId);
    if (formula) openFormulaModal(formula);
};

function openFormulaModal(formula) {
    currentModalFigure = null;
    currentModalFormula = formula;
    modalTitle.textContent = `PDF Formula #${formula.formula_id} Inspector`;
    modalSubtitle.textContent = `Page ${formula.page_number} • Tag: ${formula.formula_type || '/Formula'} • MCID: ${formula.mcids && formula.mcids.length ? formula.mcids.join(', ') : 'None'}`;
    modalImage.src = formula.image_url || '';
    modalDownloadLink.href = formula.image_url || '';
    modalDownloadLink.setAttribute('download', formula.crop_filename || `formula_${formula.formula_id}.png`);

    modalTypeLabel.textContent = 'Accessibility Math Tag';
    modalTypeDisplay.textContent = `${formula.formula_type || '/Formula'} (Structural Element • ${formula.underlying_xobjects_count || 0} XObjects)`;

    const ex = formula.excel_match;
    const effectiveAlt = (ex && ex.alt_text) || formula.alt_text || formula.actual_text || '';
    const isInjected = formula.status_label === 'Injected';

    modalAltLabel.textContent = 'Authoritative Formula /Alt Text (Editable for PDF Injection)';
    if (modalAltTextarea) {
        modalAltTextarea.value = effectiveAlt;
        modalAltTextarea.readOnly = false;
        modalAltTextarea.style.display = 'block';
    }
    if (modalAltDisplay) {
        modalAltDisplay.style.display = 'none';
    }

    if (modalInjectBtn) {
        modalInjectBtn.style.display = 'inline-flex';
        modalInjectBtn.textContent = isInjected ? '⚡ Re-Inject into PDF' : '⚡ Inject into PDF';
    }

    modalBBoxGroup.style.display = 'block';
    modalBBoxDisplay.textContent = formula.bbox ? `[${formula.bbox.join(', ')}] (${formula.bbox_width} pt × ${formula.bbox_height} pt)` : 'No bounding box';
    modalPathGroup.style.display = 'block';
    modalPathDisplay.textContent = formula.path || 'Root';
    modalMcidGroup.style.display = 'block';
    modalMcidDisplay.textContent = formula.mcids && formula.mcids.length ? `[${formula.mcids.join(', ')}]` : 'None';

    // Matched Excel preview
    if (ex) {
        modalMatchedExcelBox.style.display = 'block';
        if (modalMatchScore) modalMatchScore.style.display = 'none';
        modalExcelImage.src = ex.image_url || '';
        modalExcelRowMeta.textContent = `Row ${ex.row} (Sr. No. ${ex.sr_no || ex.row - 1}) • ${ex.filename || ''}`;
        modalExcelAltText.textContent = ex.alt_text || 'No ALT text available.';
    } else {
        modalMatchedExcelBox.style.display = 'none';
    }

    figureModal.style.display = 'flex';
}

function openExcelModal(rec) {
    currentModalFigure = null;
    modalTitle.textContent = `Excel Drawing Row ${rec.row} (Sr. ${rec.sr_no})`;
    modalSubtitle.textContent = `Manifest File: ${rec.filename || 'Drawing'} • Authoritative Alt`;
    modalImage.src = rec.image_url || '';
    modalDownloadLink.href = rec.image_url || '';
    modalDownloadLink.setAttribute('download', rec.image_filename || `excel_row_${rec.row}.png`);

    modalTypeLabel.textContent = 'Manifest Entry';
    modalTypeDisplay.textContent = `Excel Drawing Object • Row ${rec.row} • Sr. No ${rec.sr_no}`;

    modalAltLabel.textContent = 'Authoritative Client ALT Text';
    if (modalAltTextarea) {
        modalAltTextarea.value = rec.alt_text || 'No ALT text present in Excel record.';
        modalAltTextarea.readOnly = true;
        modalAltTextarea.style.display = 'block';
    }
    if (modalAltDisplay) {
        modalAltDisplay.style.display = 'none';
    }

    if (modalInjectBtn) {
        modalInjectBtn.style.display = 'none';
    }

    modalBBoxGroup.style.display = 'none';
    modalPathGroup.style.display = 'block';
    modalPathDisplay.textContent = rec.updated_alt ? `Updated Alt Text used (Original: ${rec.original_alt.substring(0, 40)}...)` : 'Original Alt Text';
    modalMcidGroup.style.display = 'none';
    modalMatchedExcelBox.style.display = 'none';

    figureModal.style.display = 'flex';
}

// Helper: Escape HTML
function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

function showLoading(title, status) {
    document.getElementById('loadingTitle').textContent = title;
    document.getElementById('loadingStatus').textContent = status;
    uploadSection.style.display = 'none';
    resultsSection.style.display = 'none';
    loadingSection.style.display = 'flex';
}

function hideLoading() {
    loadingSection.style.display = 'none';
}

document.addEventListener('DOMContentLoaded', initEvents);
