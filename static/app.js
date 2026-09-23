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

// Upload & Cancellation State
let currentUploadXHR = null;
let currentUploadId = null;
let previousVisibleSection = null;
let currentCancelHandler = null;
let currentProgressPoller = null;

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
const downloadExistingAltBtn = document.getElementById('downloadExistingAltBtn');
const downloadExistingAltBtnText = document.getElementById('downloadExistingAltBtnText');
const downloadMissingAltBtn = document.getElementById('downloadMissingAltBtn');
const downloadMissingAltBtnText = document.getElementById('downloadMissingAltBtnText');
const exportJsonBtn = document.getElementById('exportJsonBtn');
const downloadInjectedPdfBtn = document.getElementById('downloadInjectedPdfBtn');
const injectAltBtn = document.getElementById('injectAltBtn');
const injectAltBtnText = document.getElementById('injectAltBtnText');

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
const modalTypeGroup = document.getElementById('modalTypeGroup');
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
const modalSaveAltBtn = document.getElementById('modalSaveAltBtn');
const modalInjectBtn = document.getElementById('modalInjectBtn');
const modalAltTextarea = document.getElementById('modalAltTextarea');
let currentModalFigure = null;
let currentModalExcelRecord = null;

// Lightbox Modal Elements & State
const imageLightboxModal = document.getElementById('imageLightboxModal');
const lightboxTitle = document.getElementById('lightboxTitle');
const lightboxSubtitle = document.getElementById('lightboxSubtitle');
const lightboxCloseBtn = document.getElementById('lightboxCloseBtn');
const lightboxStage = document.getElementById('lightboxStage');
const zoomOutBtn = document.getElementById('zoomOutBtn');
const zoomInBtn = document.getElementById('zoomInBtn');
const zoomResetBtn = document.getElementById('zoomResetBtn');
const zoomLevelDisplay = document.getElementById('zoomLevelDisplay');
const bgToggleBtn = document.getElementById('bgToggleBtn');
const lightboxAltText = document.getElementById('lightboxAltText');
const lightboxCopyAltBtn = document.getElementById('lightboxCopyAltBtn');
const lightboxInjectBtn = document.getElementById('lightboxInjectBtn');
const lightboxDownloadLink = document.getElementById('lightboxDownloadLink');

let lightboxZoom = 1.0;
let lightboxPanX = 0;
let lightboxPanY = 0;
let isPanning = false;
let startPanX = 0;
let startPanY = 0;
let currentLightboxBgIndex = 0;
const lightboxBgModes = ['bg-grid', 'bg-white', 'bg-dark'];
let currentLightboxData = null;

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

    // Cancel Upload Button
    const cancelUploadBtn = document.getElementById('cancelUploadBtn');
    if (cancelUploadBtn) {
        cancelUploadBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (typeof currentCancelHandler === 'function') {
                currentCancelHandler();
            } else {
                cancelCurrentUpload();
            }
        });
    }

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

    // Reset / Upload Another File
    function resetToNewSession() {
        resultsSection.style.display = 'none';
        loadingSection.style.display = 'none';
        uploadSection.style.display = 'block';
        if (pdfFileInput) pdfFileInput.value = '';
        if (excelFileInput) excelFileInput.value = '';
        currentSession = null;
        currentFigures = [];
        currentFormulas = [];
        currentExcelRecords = [];
        currentMatchedFigures = [];
        selectedFigureIds.clear();
        selectedFormulaIds.clear();
        if (selectAllCheckbox) selectAllCheckbox.checked = false;
        if (tableSelectAllCheckbox) tableSelectAllCheckbox.checked = false;
        if (typeof updateSelectionUI === 'function') updateSelectionUI();
        if (downloadInjectedPdfBtn) downloadInjectedPdfBtn.style.display = 'none';
        if (tabMatchBtn) tabMatchBtn.style.display = 'none';
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    // Confirmation Modal for Uploading New File
    const confirmUploadModal = document.getElementById('confirmUploadModal');
    const confirmUploadCloseBtn = document.getElementById('confirmUploadCloseBtn');
    const confirmUploadCancelBtn = document.getElementById('confirmUploadCancelBtn');
    const confirmUploadProceedBtn = document.getElementById('confirmUploadProceedBtn');

    function openConfirmUploadModal() {
        if (confirmUploadModal) {
            confirmUploadModal.style.display = 'flex';
        }
    }

    function closeConfirmUploadModal() {
        if (confirmUploadModal) {
            confirmUploadModal.style.display = 'none';
        }
    }

    function requestUploadNewFile() {
        // Check if there is an active session or loaded figures/excel records
        const hasActiveSession = (currentSession && (currentFigures.length > 0 || currentFormulas.length > 0 || currentExcelRecords.length > 0))
            || (resultsSection && resultsSection.style.display !== 'none');
        
        if (hasActiveSession) {
            openConfirmUploadModal();
        } else {
            resetToNewSession();
        }
    }

    if (confirmUploadCloseBtn) {
        confirmUploadCloseBtn.addEventListener('click', closeConfirmUploadModal);
    }
    if (confirmUploadCancelBtn) {
        confirmUploadCancelBtn.addEventListener('click', closeConfirmUploadModal);
    }
    if (confirmUploadProceedBtn) {
        confirmUploadProceedBtn.addEventListener('click', () => {
            closeConfirmUploadModal();
            resetToNewSession();
        });
    }
    if (confirmUploadModal) {
        confirmUploadModal.addEventListener('click', (e) => {
            if (e.target === confirmUploadModal) {
                closeConfirmUploadModal();
            }
        });
    }

    if (uploadAnotherBtn) {
        uploadAnotherBtn.addEventListener('click', requestUploadNewFile);
    }

    const brandHomeBtn = document.getElementById('brandHomeBtn');
    if (brandHomeBtn) {
        brandHomeBtn.addEventListener('click', requestUploadNewFile);
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
                window.location.href = `/api/download-missing-alt-zip/${currentSession.session_id}?tab=formula`;
            } else {
                window.location.href = `/api/download-missing-alt-zip/${currentSession.session_id}?tab=pdf`;
            }
        });
    }

    if (downloadExcelZipBtn) {
        downloadExcelZipBtn.addEventListener('click', () => {
            if (currentSession && currentSession.session_id) {
                window.location.href = `/api/download-excel-zip/${currentSession.session_id}`;
            }
        });
    }

    if (exportJsonBtn) {
        exportJsonBtn.addEventListener('click', () => {
            if (currentSession && currentSession.session_id) {
                window.location.href = `/api/download-json/${currentSession.session_id}`;
            }
        });
    }

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

    const modalExpandBtn = document.getElementById('modalExpandBtn');
    if (modalExpandBtn) {
        modalExpandBtn.addEventListener('click', () => {
            if (currentModalFigure) {
                window.expandFigureSingle(currentModalFigure.figure_id);
            } else if (currentModalFormula) {
                window.expandFormulaSingle(currentModalFormula.formula_id);
            } else if (currentModalExcelRecord) {
                window.expandExcelRecord(currentModalExcelRecord.row);
            }
        });
    }

    // Download Existing Injected Alt Excel
    if (downloadExistingAltBtn) {
        downloadExistingAltBtn.addEventListener('click', () => {
            if (!currentSession || !currentSession.session_id) {
                showToast('Please upload or load a document first.');
                return;
            }
            const tabParam = currentSourceTab === 'formula' ? 'formula' : 'pdf';
            const label = currentSourceTab === 'formula' ? 'PDF Formulas & Existing Alt Text (Excel)' : 'PDF Images & Existing Alt Text (Excel)';
            showToast(`Generating ${label}...`);
            window.location.href = `/api/download-existing-alt-excel/${currentSession.session_id}?tab=${tabParam}`;
        });
    }

    // Download Missing Alt Excel
    if (downloadMissingAltBtn) {
        downloadMissingAltBtn.addEventListener('click', () => {
            if (!currentSession || !currentSession.session_id) {
                showToast('Please upload or load a document first.');
                return;
            }
            const tabParam = currentSourceTab === 'formula' ? 'formula' : 'pdf';
            showToast('Generating Missing ALT Excel template...');
            window.location.href = `/api/download-missing-alt-excel/${currentSession.session_id}?tab=${tabParam}`;
        });
    }

    // Modal Single Figure / Formula Alt Save
    if (modalSaveAltBtn) {
        modalSaveAltBtn.addEventListener('click', async () => {
            const altText = modalAltTextarea ? modalAltTextarea.value.trim() : '';
            const sid = currentSession ? currentSession.session_id : '';

            if (currentModalFormula) {
                currentModalFormula.alt_text = altText;
                currentModalFormula.has_alt = Boolean(altText);
                if (altText) {
                    currentModalFormula.status_label = 'Has /Alt';
                    selectedFormulaIds.add(currentModalFormula.formula_id);
                } else {
                    currentModalFormula.status_label = 'Missing /Alt';
                }
                const idx = currentFormulas.findIndex(f => f.formula_id === currentModalFormula.formula_id);
                if (idx !== -1) {
                    currentFormulas[idx] = { ...currentModalFormula };
                }
                if (currentSession && currentSession.formulas) {
                    const sIdx = currentSession.formulas.findIndex(f => f.formula_id === currentModalFormula.formula_id);
                    if (sIdx !== -1) currentSession.formulas[sIdx] = { ...currentModalFormula };
                }
                updateMetrics();
                refreshActiveView();

                if (sid) {
                    try {
                        await fetch(`/api/update-item-alt/${sid}`, {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                                item_type: 'formula',
                                item_id: currentModalFormula.formula_id,
                                alt_text: altText
                            })
                        });
                    } catch (err) {
                        console.error('Failed to sync saved Alt to server:', err);
                    }
                }

                modalSaveAltBtn.textContent = '✓ Saved!';
                showToast(`Formula #${currentModalFormula.formula_id} Alt Text Saved!`);
                setTimeout(() => {
                    if (modalSaveAltBtn) {
                        modalSaveAltBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Save Alt Text`;
                    }
                }, 1800);

            } else if (currentModalFigure) {
                currentModalFigure.alt_text = altText;
                currentModalFigure.has_alt = Boolean(altText);
                if (altText) {
                    currentModalFigure.status_label = 'Has /Alt';
                    selectedFigureIds.add(currentModalFigure.figure_id);
                } else {
                    currentModalFigure.status_label = 'Missing /Alt';
                }
                const idx = currentFigures.findIndex(f => f.figure_id === currentModalFigure.figure_id);
                if (idx !== -1) {
                    currentFigures[idx] = { ...currentModalFigure };
                }
                if (currentSession && currentSession.figures) {
                    const sIdx = currentSession.figures.findIndex(f => f.figure_id === currentModalFigure.figure_id);
                    if (sIdx !== -1) currentSession.figures[sIdx] = { ...currentModalFigure };
                }
                updateMetrics();
                refreshActiveView();

                if (sid) {
                    try {
                        await fetch(`/api/update-item-alt/${sid}`, {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                                item_type: 'figure',
                                item_id: currentModalFigure.figure_id,
                                alt_text: altText
                            })
                        });
                    } catch (err) {
                        console.error('Failed to sync saved Alt to server:', err);
                    }
                }

                modalSaveAltBtn.textContent = '✓ Saved!';
                showToast(`Figure #${currentModalFigure.figure_id} Alt Text Saved!`);
                setTimeout(() => {
                    if (modalSaveAltBtn) {
                        modalSaveAltBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Save Alt Text`;
                    }
                }, 1800);
            }
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
            } else if (currentSourceTab === 'match') {
                if (e.target.checked) window.selectAllMatches();
                else window.deselectAllMatches();
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
            } else if (currentSourceTab === 'match') {
                window.selectAllMatches();
            } else {
                window.selectAllFigures();
            }
        });
    }

    if (btnDeselectAll) {
        btnDeselectAll.addEventListener('click', () => {
            if (currentSourceTab === 'formula') {
                window.deselectAllFormulas();
            } else if (currentSourceTab === 'match') {
                window.deselectAllMatches();
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

    // Lightbox Modal Controls
    if (lightboxCloseBtn) {
        lightboxCloseBtn.addEventListener('click', closeLightbox);
    }
    if (imageLightboxModal) {
        imageLightboxModal.addEventListener('click', (e) => {
            if (e.target === imageLightboxModal) {
                closeLightbox();
            }
        });
    }
    if (zoomInBtn) {
        zoomInBtn.addEventListener('click', () => zoomIn(0.25));
    }
    if (zoomOutBtn) {
        zoomOutBtn.addEventListener('click', () => zoomOut(0.25));
    }
    if (zoomResetBtn) {
        zoomResetBtn.addEventListener('click', resetZoom);
    }
    if (bgToggleBtn) {
        bgToggleBtn.addEventListener('click', toggleLightboxBg);
    }
    if (lightboxCopyAltBtn) {
        lightboxCopyAltBtn.addEventListener('click', () => {
            if (currentLightboxData && currentLightboxData.altText) {
                copyToClipboard(currentLightboxData.altText, 'Copied ALT text from Lightbox!');
            } else if (lightboxAltText && lightboxAltText.textContent) {
                copyToClipboard(lightboxAltText.textContent, 'Copied ALT text from Lightbox!');
            }
        });
    }
    if (lightboxInjectBtn) {
        lightboxInjectBtn.addEventListener('click', () => {
            if (!currentLightboxData) return;
            if (currentLightboxData.itemType === 'formula') {
                window.injectAltForFormula(currentLightboxData.itemId, currentLightboxData.altText);
            } else if (currentLightboxData.itemType === 'figure') {
                window.injectAltForFigure(currentLightboxData.itemId, currentLightboxData.altText);
            }
        });
    }

    // Lightbox Panning and Wheel Listeners
    setupLightboxPanning();

    // Global Keydown Listeners for Lightbox & Modals
    window.addEventListener('keydown', (e) => {
        if (imageLightboxModal && imageLightboxModal.style.display !== 'none') {
            if (e.key === 'Escape') {
                closeLightbox();
            } else if (e.key === '+' || e.key === '=') {
                e.preventDefault();
                zoomIn(0.25);
            } else if (e.key === '-' || e.key === '_') {
                e.preventDefault();
                zoomOut(0.25);
            } else if (e.key === '0') {
                e.preventDefault();
                resetZoom();
            }
        } else if (confirmUploadModal && confirmUploadModal.style.display !== 'none') {
            if (e.key === 'Escape') {
                closeConfirmUploadModal();
            }
        } else if (figureModal && figureModal.style.display !== 'none') {
            if (e.key === 'Escape') {
                figureModal.style.display = 'none';
            }
        }
    });
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
    if (!file) return;

    const sid = (currentSession && currentSession.session_id) ? currentSession.session_id : '';
    const uploadId = 'pdf_up_' + Date.now() + '_' + Math.random().toString(36).substr(2, 7);

    showLoading(
        'Uploading and parsing PDF...',
        'Traversing StructTreeRoot and extracting /Figure tags...',
        {
            isUpload: true,
            canCancel: true,
            onCancel: cancelCurrentUpload,
            initialPercent: 0,
            initialPercentText: '0% uploaded'
        }
    );

    const formData = new FormData();
    formData.append('file', file);
    if (sid) formData.append('session_id', sid);
    formData.append('upload_id', uploadId);

    try {
        const data = await uploadWithProgress('/api/upload-pdf', formData, {
            uploadId,
            onProgress: (percent) => {
                updateLoadingProgress(
                    percent,
                    `${percent}% uploaded`,
                    'Traversing StructTreeRoot and extracting /Figure tags...'
                );
            },
            onUploadComplete: () => {
                updateLoadingProgress(
                    100,
                    '100% — Processing PDF…',
                    'Traversing StructTreeRoot and extracting /Figure tags...'
                );
            }
        });

        onPdfLoaded(data);
    } catch (err) {
        if (err.name === 'AbortError' || (err.message && err.message.toLowerCase().includes('cancel'))) {
            // Cancelled cleanly by user
            return;
        }
        alert('Error: ' + err.message);
        hideLoading();
        if (previousVisibleSection === 'results' && currentSession) {
            resultsSection.style.display = 'block';
        } else {
            uploadSection.style.display = 'block';
        }
    }
}

async function handleLoadSamplePdf() {
    showLoading('Loading Chapter 15 Sample PDF...', 'Extracting 88 /Figure tags and marked content coordinates...', {
        isUpload: false,
        canCancel: false
    });
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
    if (!file) return;

    const sid = currentSession ? currentSession.session_id : '';
    const uploadId = 'excel_up_' + Date.now() + '_' + Math.random().toString(36).substr(2, 7);

    if (currentProgressPoller) {
        clearInterval(currentProgressPoller);
        currentProgressPoller = null;
    }

    showLoading(
        'Uploading Excel ALT Manifest…',
        'Extracting drawings, filenames, and authoritative ALT text...',
        {
            isUpload: true,
            canCancel: true,
            onCancel: cancelCurrentUpload,
            initialPercent: 0,
            initialPercentText: '0% uploaded'
        }
    );

    const formData = new FormData();
    formData.append('file', file);
    if (sid) formData.append('session_id', sid);
    formData.append('upload_id', uploadId);

    // Live background progress poller
    const startProgressPolling = () => {
        if (currentProgressPoller) return;
        currentProgressPoller = setInterval(async () => {
            try {
                const res = await fetch(`/api/upload-progress/${uploadId}`);
                if (res.ok) {
                    const prog = await res.json();
                    if (prog && prog.percent > 0) {
                        const titleText = prog.title || 'Processing Excel…';
                        const statusMsg = prog.status || 'Extracting visual records...';
                        updateLoadingProgress(
                            prog.percent,
                            `${prog.percent}% — ${titleText}`,
                            statusMsg
                        );
                    }
                }
            } catch (e) {
                // Ignore transient polling error
            }
        }, 120);
    };

    try {
        const uploadPromise = uploadWithProgress('/api/upload-excel', formData, {
            uploadId,
            onProgress: (percent) => {
                if (percent < 100) {
                    updateLoadingProgress(
                        percent,
                        `${percent}% uploaded`,
                        'Uploading Excel ALT manifest file to server...'
                    );
                } else {
                    updateLoadingProgress(
                        10,
                        '10% — Processing Excel…',
                        'Scanning worksheet and embedded DrawingML images...'
                    );
                    startProgressPolling();
                }
            },
            onUploadComplete: () => {
                updateLoadingProgress(
                    10,
                    '10% — Processing Excel…',
                    'Scanning worksheet and embedded DrawingML images...'
                );
                startProgressPolling();
            }
        });

        // Start polling shortly after sending in case upload finished in milliseconds on localhost
        setTimeout(startProgressPolling, 150);

        const data = await uploadPromise;
        if (currentProgressPoller) {
            clearInterval(currentProgressPoller);
            currentProgressPoller = null;
        }

        updateLoadingProgress(100, '100% — Processing Complete', 'Finalizing manifest gallery...');
        setTimeout(() => {
            onExcelLoaded(data);
        }, 200);
    } catch (err) {
        if (currentProgressPoller) {
            clearInterval(currentProgressPoller);
            currentProgressPoller = null;
        }
        if (err.name === 'AbortError' || (err.message && err.message.toLowerCase().includes('cancel'))) {
            // Cancelled cleanly by user
            return;
        }
        alert('Error loading Excel: ' + err.message);
        hideLoading();
        if (previousVisibleSection === 'results' && currentSession) {
            resultsSection.style.display = 'block';
        } else {
            uploadSection.style.display = 'block';
        }
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
    if (downloadExistingAltBtn) {
        downloadExistingAltBtn.style.display = 'inline-flex';
        if (downloadExistingAltBtnText) downloadExistingAltBtnText.textContent = 'Download Image & Existing Alt Text';
    }
    if (downloadMissingAltBtn) {
        downloadMissingAltBtn.style.display = 'inline-flex';
        if (downloadMissingAltBtnText) downloadMissingAltBtnText.textContent = 'Download Missing Alt (Excel)';
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
        if (data.has_alt_count !== undefined) currentSession.has_alt_count = data.has_alt_count;
        if (data.missing_alt_count !== undefined) currentSession.missing_alt_count = data.missing_alt_count;
        if (data.has_formula_alt_count !== undefined) currentSession.has_formula_alt_count = data.has_formula_alt_count;
        if (data.missing_formula_alt_count !== undefined) currentSession.missing_formula_alt_count = data.missing_formula_alt_count;
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

    // Refresh all figures/formulas/tabs metrics and selection counts
    updateMetrics();
    updateSelectionUI();

    // If PDF figures exist, keep user on PDF Figures tab or current active tab with refreshed matches
    const targetTab = currentSourceTab === 'formula' ? 'formula' : (currentFigures.length > 0 ? 'pdf' : 'excel');
    switchSourceTab(targetTab);
    showToast(`Loaded ${currentExcelRecords.length.toLocaleString()} Excel records! Matched Alt text linked to PDF figures.`);
}

function updateTabVisibility() {
    const hasPdf = currentFigures && currentFigures.length > 0;
    const hasFormulas = currentFormulas && currentFormulas.length > 0;
    const hasExcel = currentExcelRecords && currentExcelRecords.length > 0;

    tabPdfBtn.style.display = hasPdf ? 'inline-flex' : 'none';
    if (tabFormulaBtn) tabFormulaBtn.style.display = hasFormulas ? 'inline-flex' : 'none';
    tabExcelBtn.style.display = hasExcel ? 'inline-flex' : 'none';

    if ((hasPdf || hasFormulas) && hasExcel) {
        tabMatchBtn.style.display = 'inline-flex';
        const matchedFigures = (currentFigures || []).filter(f => f.excel_match || f.status_label === 'Injected');
        const matchedFormulas = (currentFormulas || []).filter(f => f.excel_match || f.status_label === 'Injected');
        const allMatched = [...matchedFigures, ...matchedFormulas];
        const matchedCount = allMatched.length;
        const totalPdfItems = (currentFigures ? currentFigures.length : 0) + (currentFormulas ? currentFormulas.length : 0);

        tabMatchBadge.textContent = matchedCount.toLocaleString();
        statMatchPdfTotal.textContent = totalPdfItems.toLocaleString();
        statMatchMatched.textContent = matchedCount.toLocaleString();
        statMatchReady.textContent = matchedCount.toLocaleString();

        if (matchedCount > 0) {
            const totalConfidence = allMatched.reduce((sum, item) => {
                const conf = (typeof item.confidence === 'number' && !isNaN(item.confidence)) ? item.confidence : (item.excel_match ? 0.95 : 0.85);
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
    if (excelMetricsGrid) excelMetricsGrid.style.display = tab === 'excel' ? 'grid' : 'none';
    if (matchMetricsGrid) matchMetricsGrid.style.display = tab === 'match' ? 'grid' : 'none';

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
        if (downloadExcelZipBtn) downloadExcelZipBtn.style.display = 'none';
        if (downloadExistingAltBtn) downloadExistingAltBtn.style.display = 'none';
        if (downloadMissingAltBtn) downloadMissingAltBtn.style.display = 'none';
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
        if (downloadFormulasZipBtn) downloadFormulasZipBtn.style.display = 'none';
        if (downloadExcelZipBtn) downloadExcelZipBtn.style.display = 'none';
        if (downloadExistingAltBtn) {
            downloadExistingAltBtn.style.display = 'inline-flex';
            if (downloadExistingAltBtnText) downloadExistingAltBtnText.textContent = 'Download Formula & Existing Alt Text';
        }
        if (downloadMissingAltBtn) {
            downloadMissingAltBtn.style.display = 'inline-flex';
            if (downloadMissingAltBtnText) downloadMissingAltBtnText.textContent = 'Download Missing Formulas (Excel)';
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
        if (downloadExcelZipBtn) downloadExcelZipBtn.style.display = 'none';
        if (downloadExistingAltBtn) {
            downloadExistingAltBtn.style.display = 'inline-flex';
            if (downloadExistingAltBtnText) downloadExistingAltBtnText.textContent = 'Download Image & Existing Alt Text';
        }
        if (downloadMissingAltBtn) {
            downloadMissingAltBtn.style.display = 'inline-flex';
            if (downloadMissingAltBtnText) downloadMissingAltBtnText.textContent = 'Download Missing Alt (Excel)';
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
        docMeta.textContent = 'Automated Visual & Mathematical Alignment of PDF Figures & Formulas to Authoritative Excel Sources';
        if (downloadUnselectedLabel) downloadUnselectedLabel.textContent = 'Download Matches (ZIP)';
        if (downloadFormulasZipBtn) downloadFormulasZipBtn.style.display = 'none';
        if (downloadExcelZipBtn) downloadExcelZipBtn.style.display = 'none';
        if (downloadExistingAltBtn) downloadExistingAltBtn.style.display = 'none';
        if (downloadMissingAltBtn) downloadMissingAltBtn.style.display = 'none';
        if (selectionBar) selectionBar.style.display = 'flex';
        filterHasImgBtn.style.display = 'none';
        searchInput.placeholder = 'Search matched figures & formulas by page, ID, MCID, or Alt text...';
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
// METRICS & STATS SYNCHRONIZATION
// ==========================================
function updateMetrics() {
    if (currentFigures) {
        const hasAltCount = currentFigures.filter(f => Boolean(f.has_alt || f.alt_text || (f.excel_match && f.excel_match.alt_text))).length;
        const missingCount = currentFigures.length - hasAltCount;
        if (statFiguresCount) statFiguresCount.textContent = currentFigures.length;
        if (statHasAlt) statHasAlt.textContent = hasAltCount;
        if (statMissingAlt) statMissingAlt.textContent = missingCount;
        if (tabPdfBadge) tabPdfBadge.textContent = currentFigures.length;
    }
    if (currentFormulas) {
        const hasFormulaAltCount = currentFormulas.filter(f => Boolean(f.has_alt || f.alt_text || (f.excel_match && f.excel_match.alt_text) || f.actual_text)).length;
        const missingFormulaCount = currentFormulas.length - hasFormulaAltCount;
        if (statFormulasCount) statFormulasCount.textContent = currentFormulas.length.toLocaleString();
        if (statFormulaHasAlt) statFormulaHasAlt.textContent = hasFormulaAltCount.toLocaleString();
        if (statFormulaMissingAlt) statFormulaMissingAlt.textContent = missingFormulaCount.toLocaleString();
        if (tabFormulaBadge) tabFormulaBadge.textContent = currentFormulas.length.toLocaleString();
    }
    if (currentExcelRecords) {
        const hasAltCount = currentExcelRecords.filter(r => r.has_alt).length;
        const missingCount = currentExcelRecords.length - hasAltCount;
        if (statExcelTotal) statExcelTotal.textContent = currentExcelRecords.length.toLocaleString();
        if (statExcelHasAlt) statExcelHasAlt.textContent = hasAltCount.toLocaleString();
        if (statExcelMissingAlt) statExcelMissingAlt.textContent = missingCount.toLocaleString();
        if (tabExcelBadge) tabExcelBadge.textContent = currentExcelRecords.length.toLocaleString();
    }
    updateTabVisibility();
    updateFilterCounts();
}

// ==========================================
// FILTER COUNTS CALCULATION
// ==========================================
function updateFilterCounts() {
    if (currentSourceTab === 'pdf') {
        const total = currentFigures.length;
        const hasAlt = currentFigures.filter(f => Boolean(f.has_alt || f.alt_text || (f.excel_match && f.excel_match.alt_text))).length;
        const missing = total - hasAlt;

        countAll.textContent = total;
        countHasAlt.textContent = hasAlt;
        countMissing.textContent = missing;
    } else if (currentSourceTab === 'formula') {
        const total = currentFormulas.length;
        const hasAlt = currentFormulas.filter(f => Boolean(f.has_alt || f.alt_text || (f.excel_match && f.excel_match.alt_text) || f.actual_text)).length;
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
        const matchedFigures = (currentFigures || []).filter(f => f.excel_match || f.status_label === 'Injected');
        const matchedFormulas = (currentFormulas || []).filter(f => f.excel_match || f.status_label === 'Injected');
        const totalMatched = matchedFigures.length + matchedFormulas.length;

        countAll.textContent = totalMatched.toLocaleString();
        countHasAlt.textContent = matchedFigures.length.toLocaleString();
        countMissing.textContent = matchedFormulas.length.toLocaleString();

        if (filterAllBtn) filterAllBtn.innerHTML = `All Matches (<span id="countAll">${totalMatched.toLocaleString()}</span>)`;
        if (filterHasAltBtn) filterHasAltBtn.innerHTML = `Figures (<span id="countHasAlt">${matchedFigures.length.toLocaleString()}</span>)`;
        if (filterMissingBtn) filterMissingBtn.innerHTML = `Formulas (<span id="countMissing">${matchedFormulas.length.toLocaleString()}</span>)`;
        return;
    }

    if (filterAllBtn) filterAllBtn.innerHTML = `All Items (<span id="countAll">${countAll.textContent}</span>)`;
    if (filterHasAltBtn) filterHasAltBtn.innerHTML = `Has Alt (<span id="countHasAlt">${countHasAlt.textContent}</span>)`;
    if (filterMissingBtn) filterMissingBtn.innerHTML = `Missing Alt (<span id="countMissing">${countMissing.textContent}</span>)`;
}

// ==========================================
// PDF FIGURES RENDERING
// ==========================================
function renderPdfFigures() {
    const q = searchInput.value.toLowerCase().trim();

    const filtered = currentFigures.filter(fig => {
        const hasAlt = Boolean(fig.has_alt || fig.alt_text || (fig.excel_match && fig.excel_match.alt_text));
        if (activeFilter === 'missing' && hasAlt) return false;
        if (activeFilter === 'has_alt' && !hasAlt) return false;

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
                    <div class="card-dual-image-grid" onclick="event.stopPropagation(); window.expandFigureDual(${fig.figure_id});" title="Click to open Side-by-Side Zoom & Compare">
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
                    <div class="compare-action-row">
                        <button class="btn-compare-expand" onclick="event.stopPropagation(); window.expandFigureDual(${fig.figure_id})" title="Side-by-Side Zoom & Compare">
                            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/></svg>
                            <span>Side-by-Side Zoom</span>
                        </button>
                    </div>
                `;
            } else {
                imagesHtml = `
                    <div class="figure-image-wrapper" onclick="event.stopPropagation(); window.expandFigureSingle(${fig.figure_id});" style="cursor: pointer;" title="Click to Expand & Zoom">
                        <span class="page-chip">Page ${fig.page_number || '?'}</span>
                        <span class="mcid-chip">${mcidText}</span>
                        <img src="${fig.image_url || ''}" alt="Figure ${fig.figure_id}" loading="lazy">
                        <button class="btn-img-expand" onclick="event.stopPropagation(); window.expandFigureSingle(${fig.figure_id})" title="Expand & Zoom Figure">
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/></svg>
                            <span>Expand</span>
                        </button>
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
                            <span>${isInjected ? '✓ INJECTED STRUCTTREE ALT:' : (ex ? 'AUTHORITATIVE ALT TEXT FROM EXCEL:' : (fig.alt_text ? 'SAVED ALT TEXT (READY TO INJECT):' : 'CURRENT PDF /ALT:'))}</span>
                        </div>
                        <div class="alt-body-text">${escapeHtml(effectiveAlt || 'No /Alt accessibility text defined in StructTree.')}</div>
                    </div>
                    <div class="card-footer-row">
                        <span class="footer-hint">${ex ? `Row ${ex.row} • Sr. ${ex.sr_no}` : `Page ${fig.page_number}`}</span>
                        <div style="display:flex; gap:6px;">
                            ${hasEffectiveAlt ? `
                            <button class="btn btn-copy-alt btn-sm" onclick="event.stopPropagation(); window.copyFigureAlt(${fig.figure_id})" style="padding: 5px 10px; font-size: 0.78rem;" title="Copy Alt Text">
                                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                                <span>Copy</span>
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
                    <button class="btn btn-copy-alt btn-sm" onclick="event.stopPropagation(); window.copyFigureAlt(${fig.figure_id})" style="padding: 4px 8px; font-size: 0.75rem;" title="Copy Alt Text">
                        Copy
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
                        <div class="card-dual-image-grid" onclick="event.stopPropagation(); window.expandFormulaDual(${formula.formula_id});" title="Click to open Side-by-Side Zoom & Compare">
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
                        <div class="compare-action-row">
                            <button class="btn-compare-expand" onclick="event.stopPropagation(); window.expandFormulaDual(${formula.formula_id})" title="Side-by-Side Zoom & Compare">
                                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/></svg>
                                <span>Side-by-Side Zoom</span>
                            </button>
                        </div>
                    `;
                } else {
                    imagesHtml = `
                        <div class="figure-image-wrapper" onclick="event.stopPropagation(); window.expandFormulaSingle(${formula.formula_id});" style="cursor: pointer;" title="Click to Expand & Zoom">
                            <span class="page-chip">Page ${formula.page_number || '?'}</span>
                            <span class="mcid-chip">${mcidText}</span>
                            <img src="${formula.image_url || ''}" alt="Formula ${formula.formula_id}" loading="lazy">
                            <button class="btn-img-expand" onclick="event.stopPropagation(); window.expandFormulaSingle(${formula.formula_id})" title="Expand & Zoom Formula">
                                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/></svg>
                                <span>Expand</span>
                            </button>
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
                                <span>${isInjected ? '✓ INJECTED FORMULA ALT:' : (ex ? 'AUTHORITATIVE ALT TEXT FROM EXCEL:' : (formula.alt_text ? 'SAVED FORMULA ALT (READY TO INJECT):' : 'CURRENT /ALT OR /ACTUALTEXT:'))}</span>
                            </div>
                            <div class="alt-body-text">${escapeHtml(effectiveAlt || 'No /Alt or /ActualText defined in StructTree.')}</div>
                        </div>
                        <div class="card-footer-row">
                            <span class="footer-hint">${ex ? `Row ${ex.row} • Sr. ${ex.sr_no}` : `Page ${formula.page_number}`}</span>
                            <div style="display:flex; gap:6px;">
                                ${hasEffectiveAlt ? `
                                <button class="btn btn-copy-alt btn-sm" onclick="event.stopPropagation(); window.copyFormulaAlt(${formula.formula_id})" style="padding: 5px 10px; font-size: 0.78rem;" title="Copy Alt Text">
                                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                                    <span>Copy</span>
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
                        <button class="btn btn-copy-alt btn-sm" onclick="event.stopPropagation(); window.copyFormulaAlt(${formula.formula_id})" style="padding: 4px 8px; font-size: 0.75rem;" title="Copy Alt Text">
                            Copy
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
                    ${hasImg ? `
                    <button class="btn-img-expand" onclick="event.stopPropagation(); window.expandExcelRecord(${rec.row})" title="Expand & Zoom Image">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/></svg>
                        <span>Expand</span>
                    </button>` : ''}
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
window.copyExcelAlt = function (row) {
    const rec = currentExcelRecords.find(r => r.row === row);
    if (rec && rec.alt_text) {
        copyToClipboard(rec.alt_text, `Copied Row ${row} ALT text!`);
    } else {
        showToast(`Row ${row} has no ALT text.`);
    }
};

window.copyFigureAlt = function (figId) {
    const fig = currentFigures.find(f => f.figure_id === figId);
    if (!fig) return;
    const text = (fig.excel_match && fig.excel_match.alt_text) || fig.alt_text;
    if (text) {
        copyToClipboard(text, `Copied Alt for Figure ${figId}!`);
    } else {
        showToast(`Figure ${figId} has no ALT text.`);
    }
};

window.copyFormulaAlt = function (formulaId) {
    const formula = currentFormulas.find(f => f.formula_id === formulaId);
    if (!formula) return;
    const text = (formula.excel_match && formula.excel_match.alt_text) || formula.alt_text || formula.actual_text;
    if (text) {
        copyToClipboard(text, `Copied Alt for Formula #${formulaId}!`);
    } else {
        showToast(`Formula #${formulaId} has no ALT text.`);
    }
};

window.openExcelModalByRow = function (row) {
    const rec = currentExcelRecords.find(r => r.row === row);
    if (rec) openExcelModal(rec);
};

// ==========================================
// ALT TEXT INJECTION FUNCTIONS
// ==========================================
window.injectAltForFigure = async function (figId, customAlt) {
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

window.injectAltForFormula = async function (formulaId, customAlt) {
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
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><path d="M9 15l2 2 4-4"></path></svg>
                <span>✓ Injected All (${data.injected_count})</span>
            `;
            injectAltBtn.disabled = false;

            showToast(`Successfully injected ${data.injected_count} ALT texts into StructTreeRoot /Formula tags!`);
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
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><path d="M9 15l2 2 4-4"></path></svg>
            <span>✓ Injected All (${data.injected_count})</span>
        `;
        injectAltBtn.disabled = false;

        showToast(`Successfully injected ${data.injected_count} ALT texts into StructTreeRoot /Figure tags!`);
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
function setInjectSelectedButtonState(state, count = 0) {
    if (!injectSelectedBtn) return;
    if (state === 'loading' || state === true) {
        injectSelectedBtn.innerHTML = `
            <div class="spinner" style="width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; display:inline-block; vertical-align:middle; margin-right:6px;"></div>
            <span>Injecting (${count} Selected)...</span>
        `;
        injectSelectedBtn.disabled = true;
    } else if (state === 'success') {
        injectSelectedBtn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
                <polyline points="20 6 9 17 4 12"></polyline>
            </svg>
            <span id="injectSelectedBtnText">✓ Injected (${count} Selected)</span>
        `;
        injectSelectedBtn.disabled = false;
    } else {
        injectSelectedBtn.innerHTML = `
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <path d="M9 15l2 2 4-4"></path>
            </svg>
            <span id="injectSelectedBtnText">Inject Selected Alt into PDF (${count})</span>
        `;
        injectSelectedBtn.disabled = count === 0;
    }
}

function setRemoveAltButtonState(state, count = 0, isFormula = false) {
    if (!removeAltBtn) return;
    if (state === 'loading' || state === true) {
        removeAltBtn.innerHTML = `
            <div class="spinner" style="width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; display:inline-block; vertical-align:middle; margin-right:6px;"></div>
            <span>${isFormula ? 'Removing Formula Alt...' : 'Removing Alt...'}</span>
        `;
        removeAltBtn.disabled = true;
    } else {
        const text = count > 0
            ? (isFormula ? `Remove Selected Formula Alt (${count})` : `Remove Selected Alt (${count})`)
            : (isFormula ? `Remove All Formula Alt` : `Remove All Alt Text`);
        removeAltBtn.innerHTML = `
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
                <polyline points="3 6 5 6 21 6"></polyline>
                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                <line x1="10" y1="11" x2="10" y2="17"></line>
                <line x1="14" y1="11" x2="14" y2="17"></line>
            </svg>
            <span id="removeAltBtnText">${text}</span>
        `;
        removeAltBtn.disabled = false;
    }
}

function updateSelectionUI() {
    if (currentSourceTab === 'formula') {
        const selectableFormulas = currentFormulas.filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text || f.actual_text);
        const totalSelectable = selectableFormulas.length;
        const selectedCount = selectedFormulaIds.size;

        if (selectedCountBadge) selectedCountBadge.textContent = selectedCount;
        if (totalSelectableBadge) totalSelectableBadge.textContent = totalSelectable;

        setInjectSelectedButtonState('ready', selectedCount);
        setRemoveAltButtonState('ready', selectedCount, true);

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

    if (currentSourceTab === 'match') {
        const selectableFigures = (currentFigures || []).filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text);
        const selectableFormulas = (currentFormulas || []).filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text || f.actual_text);
        const totalSelectable = selectableFigures.length + selectableFormulas.length;
        const selectedCount = selectedFigureIds.size + selectedFormulaIds.size;

        if (selectedCountBadge) selectedCountBadge.textContent = selectedCount;
        if (totalSelectableBadge) totalSelectableBadge.textContent = totalSelectable;

        setInjectSelectedButtonState('ready', selectedCount);
        setRemoveAltButtonState('ready', selectedCount, false);

        const allChecked = totalSelectable > 0 && selectedCount === totalSelectable;
        const isIndeterminate = selectedCount > 0 && selectedCount < totalSelectable;

        if (selectAllCheckbox) {
            selectAllCheckbox.checked = allChecked;
            selectAllCheckbox.indeterminate = isIndeterminate;
        }

        // Sync match cards in DOM for figures
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

        // Sync match cards in DOM for formulas
        document.querySelectorAll('.match-card[data-formula-id]').forEach(card => {
            const formulaId = parseInt(card.getAttribute('data-formula-id'), 10);
            const isSel = selectedFormulaIds.has(formulaId);
            card.classList.toggle('selected', isSel);

            const cb = card.querySelector('.match-card-checkbox');
            if (cb) cb.checked = isSel;

            const btn = card.querySelector('.btn-card-select-toggle');
            if (btn) {
                btn.classList.toggle('selected', isSel);
                btn.textContent = isSel ? '✓ Selected' : '+ Select';
            }
        });
        return;
    }

    const selectableFigures = currentFigures.filter(f => (f.excel_match && f.excel_match.alt_text) || f.alt_text);
    const totalSelectable = selectableFigures.length;
    const selectedCount = selectedFigureIds.size;

    if (selectedCountBadge) selectedCountBadge.textContent = selectedCount;
    if (totalSelectableBadge) totalSelectableBadge.textContent = totalSelectable;

    setInjectSelectedButtonState('ready', selectedCount);
    setRemoveAltButtonState('ready', selectedCount, false);

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

window.toggleFigureSelection = function (figId, forceState) {
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

window.selectAllFigures = function () {
    currentFigures.forEach(f => {
        const hasAlt = (f.excel_match && f.excel_match.alt_text) || f.alt_text;
        if (hasAlt) {
            selectedFigureIds.add(f.figure_id);
        }
    });
    updateSelectionUI();
    showToast(`Selected all ${selectedFigureIds.size} figure boxes for injection.`);
};

window.deselectAllFigures = function () {
    selectedFigureIds.clear();
    updateSelectionUI();
    showToast('Deselected all figures.');
};

window.toggleFormulaSelection = function (formulaId, forceState) {
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

window.selectAllFormulas = function () {
    currentFormulas.forEach(f => {
        const hasAlt = (f.excel_match && f.excel_match.alt_text) || f.alt_text || f.actual_text;
        if (hasAlt) {
            selectedFormulaIds.add(f.formula_id);
        }
    });
    updateSelectionUI();
    showToast(`Selected all ${selectedFormulaIds.size} formula boxes for injection.`);
};

window.deselectAllFormulas = function () {
    selectedFormulaIds.clear();
    updateSelectionUI();
    showToast('Deselected all formulas.');
};

window.selectAllMatches = function () {
    (currentFigures || []).forEach(f => {
        const hasAlt = (f.excel_match && f.excel_match.alt_text) || f.alt_text;
        if (hasAlt) {
            selectedFigureIds.add(f.figure_id);
        }
    });
    (currentFormulas || []).forEach(f => {
        const hasAlt = (f.excel_match && f.excel_match.alt_text) || f.alt_text || f.actual_text;
        if (hasAlt) {
            selectedFormulaIds.add(f.formula_id);
        }
    });
    updateSelectionUI();
    const total = selectedFigureIds.size + selectedFormulaIds.size;
    showToast(`Selected all ${total} matched boxes for injection.`);
};

window.deselectAllMatches = function () {
    selectedFigureIds.clear();
    selectedFormulaIds.clear();
    updateSelectionUI();
    showToast('Deselected all matched items.');
};

async function handleInjectSelected() {
    if (!currentSession || !currentSession.session_id) {
        showToast('Please load or upload a PDF first.');
        return;
    }

    if (currentSourceTab === 'match') {
        const totalSelected = selectedFigureIds.size + selectedFormulaIds.size;
        if (totalSelected === 0) {
            showToast('Please select at least one matched figure or formula box to inject.');
            return;
        }

        const figInjections = {};
        selectedFigureIds.forEach(id => {
            const fig = currentFigures.find(f => f.figure_id === id);
            if (fig) {
                const alt = (fig.excel_match && fig.excel_match.alt_text) || fig.alt_text;
                if (alt) figInjections[id] = alt.trim();
            }
        });

        const formInjections = {};
        selectedFormulaIds.forEach(id => {
            const formula = currentFormulas.find(f => f.formula_id === id);
            if (formula) {
                const alt = (formula.excel_match && formula.excel_match.alt_text) || formula.alt_text || formula.actual_text;
                if (alt) formInjections[id] = alt.trim();
            }
        });

        if (Object.keys(figInjections).length === 0 && Object.keys(formInjections).length === 0) {
            showToast('None of the selected items have ALT text available.');
            return;
        }

        setInjectSelectedButtonState('loading', totalSelected);

        try {
            let injectedFigCount = 0;
            let injectedFormCount = 0;

            if (Object.keys(figInjections).length > 0) {
                const resFig = await fetch(`/api/inject-alt/${currentSession.session_id}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ injections: figInjections })
                });
                if (resFig.ok) {
                    const dataFig = await resFig.json();
                    currentFigures = dataFig.figures;
                    if (currentSession) {
                        currentSession.figures = dataFig.figures;
                        currentSession.has_alt_count = dataFig.has_alt_count;
                        currentSession.missing_alt_count = dataFig.missing_alt_count;
                    }
                    injectedFigCount = dataFig.injected_count || Object.keys(figInjections).length;
                }
            }

            if (Object.keys(formInjections).length > 0) {
                const resForm = await fetch(`/api/inject-formula-alt/${currentSession.session_id}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ injections: formInjections })
                });
                if (resForm.ok) {
                    const dataForm = await resForm.json();
                    currentFormulas = dataForm.formulas;
                    if (currentSession) {
                        currentSession.formulas = dataForm.formulas;
                        currentSession.has_formula_alt_count = dataForm.has_formula_alt_count;
                        currentSession.missing_formula_alt_count = dataForm.missing_formula_alt_count;
                    }
                    injectedFormCount = dataForm.injected_count || Object.keys(formInjections).length;
                }
            }

            if (currentSession) {
                currentSession.has_injected_pdf = true;
            }

            if (downloadInjectedPdfBtn) {
                downloadInjectedPdfBtn.href = `/api/download-injected-pdf/${currentSession.session_id}`;
                downloadInjectedPdfBtn.style.display = 'inline-flex';
            }

            updateMetrics();
            refreshActiveView();
            setInjectSelectedButtonState('success', injectedFigCount + injectedFormCount);
            showToast(`✓ Injected ALT text for ${injectedFigCount} figure(s) and ${injectedFormCount} formula(s)!`);

        } catch (err) {
            alert('Injection failed: ' + err.message);
            updateSelectionUI();
        }
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

        const countToInject = selectedFormulaIds.size;
        setInjectSelectedButtonState('loading', countToInject);

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

            showToast(`Successfully injected ${data.injected_count} selected formula ALT texts into StructTreeRoot!`);
            refreshActiveView();
            setInjectSelectedButtonState('success', data.injected_count || countToInject);

        } catch (err) {
            alert('Formula injection failed: ' + err.message);
            updateSelectionUI();
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

    const countToInject = selectedFigureIds.size;
    setInjectSelectedButtonState('loading', countToInject);

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

        showToast(`Successfully injected ${data.injected_count} selected ALT texts into StructTreeRoot!`);
        refreshActiveView();
        setInjectSelectedButtonState('success', data.injected_count || countToInject);

    } catch (err) {
        alert('Injection failed: ' + err.message);
        updateSelectionUI();
    }
}

async function handleRemoveAlt() {
    if (!currentSession || !currentSession.session_id) {
        showToast('Please load or upload a PDF first.');
        return;
    }

    if (currentSourceTab === 'match') {
        const selectedCount = selectedFigureIds.size + selectedFormulaIds.size;
        const isTargeted = selectedCount > 0;
        const confirmMsg = isTargeted
            ? `Are you sure you want to remove /Alt text from the ${selectedCount} selected item(s) in the PDF?`
            : `Are you sure you want to remove /Alt text from ALL matched figures & formulas in the PDF?`;

        if (!confirm(confirmMsg)) return;

        setRemoveAltButtonState('loading', selectedCount, false);

        try {
            if (isTargeted) {
                if (selectedFigureIds.size > 0) {
                    const resFig = await fetch(`/api/remove-alt/${currentSession.session_id}`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ figure_ids: Array.from(selectedFigureIds) })
                    });
                    if (resFig.ok) {
                        const dataFig = await resFig.json();
                        currentFigures = dataFig.figures;
                        if (currentSession) currentSession.figures = dataFig.figures;
                    }
                }
                if (selectedFormulaIds.size > 0) {
                    const resForm = await fetch(`/api/remove-formula-alt/${currentSession.session_id}`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ formula_ids: Array.from(selectedFormulaIds) })
                    });
                    if (resForm.ok) {
                        const dataForm = await resForm.json();
                        currentFormulas = dataForm.formulas;
                        if (currentSession) currentSession.formulas = dataForm.formulas;
                    }
                }
            } else {
                const resFig = await fetch(`/api/remove-alt/${currentSession.session_id}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ figure_ids: null })
                });
                if (resFig.ok) {
                    const dataFig = await resFig.json();
                    currentFigures = dataFig.figures;
                    if (currentSession) currentSession.figures = dataFig.figures;
                }
                const resForm = await fetch(`/api/remove-formula-alt/${currentSession.session_id}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ formula_ids: null })
                });
                if (resForm.ok) {
                    const dataForm = await resForm.json();
                    currentFormulas = dataForm.formulas;
                    if (currentSession) currentSession.formulas = dataForm.formulas;
                }
            }

            updateMetrics();
            refreshActiveView();
            showToast('🗑️ Successfully removed /Alt text from matched items in the PDF!');
        } catch (err) {
            alert('Removal failed: ' + err.message);
        } finally {
            updateSelectionUI();
        }
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

        setRemoveAltButtonState('loading', selectedCount, true);

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

        } catch (err) {
            alert('Formula removal failed: ' + err.message);
        } finally {
            updateSelectionUI();
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

    setRemoveAltButtonState('loading', selectedCount, false);

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

    } catch (err) {
        alert('Removal failed: ' + err.message);
    } finally {
        updateSelectionUI();
    }
}

// ==========================================
// MATCHED SIDE-BY-SIDE VIEW RENDERING
// ==========================================
function renderMatched() {
    matchGrid.innerHTML = '';
    const q = searchInput.value.toLowerCase().trim();

    const matchedFigures = (currentFigures || []).filter(f => f.excel_match || f.status_label === 'Injected').map(f => ({
        ...f,
        match_category: 'figure',
        sort_id: f.figure_id
    }));

    const matchedFormulas = (currentFormulas || []).filter(f => f.excel_match || f.status_label === 'Injected').map(f => ({
        ...f,
        match_category: 'formula',
        sort_id: f.formula_id
    }));

    let allMatched = [...matchedFigures, ...matchedFormulas];

    // Filter by type if user clicked filter buttons
    if (activeFilter === 'has_alt') {
        allMatched = allMatched.filter(item => item.match_category === 'figure');
    } else if (activeFilter === 'missing') {
        allMatched = allMatched.filter(item => item.match_category === 'formula');
    }

    if (q) {
        allMatched = allMatched.filter(item => {
            const pageMatch = String(item.page_number || '').includes(q);
            const idMatch = String(item.figure_id || item.formula_id || '').includes(q);
            const mcidMatch = (item.mcids || []).some(m => String(m).includes(q));
            const altMatch = (item.alt_text || '').toLowerCase().includes(q);
            const actualMatch = (item.actual_text || '').toLowerCase().includes(q);
            const ex = item.excel_match;
            const exAltMatch = ex && (ex.alt_text || '').toLowerCase().includes(q);
            const exFnMatch = ex && (ex.filename || '').toLowerCase().includes(q);
            const exRowMatch = ex && String(ex.row || '').includes(q);
            const exSrMatch = ex && String(ex.sr_no || '').includes(q);
            const typeMatch = item.match_category.includes(q) || (item.formula_type || '').toLowerCase().includes(q);
            return pageMatch || idMatch || mcidMatch || altMatch || actualMatch || exAltMatch || exFnMatch || exRowMatch || exSrMatch || typeMatch;
        });
    }

    // Sort sequentially by PDF page number, then vertical coordinate (top to bottom), then horizontal, then ID
    allMatched.sort((a, b) => {
        const pA = a.page_number || 0;
        const pB = b.page_number || 0;
        if (pA !== pB) return pA - pB;
        const yA = (a.bbox && a.bbox.length > 1) ? a.bbox[1] : 0;
        const yB = (b.bbox && b.bbox.length > 1) ? b.bbox[1] : 0;
        if (Math.abs(yA - yB) > 2) return yA - yB;
        const xA = (a.bbox && a.bbox.length > 0) ? a.bbox[0] : 0;
        const xB = (b.bbox && b.bbox.length > 0) ? b.bbox[0] : 0;
        if (Math.abs(xA - xB) > 2) return xA - xB;
        return a.sort_id - b.sort_id;
    });

    if (allMatched.length === 0) {
        matchGrid.innerHTML = `
            <div style="text-align:center; padding:60px 20px; color:var(--text-dim);">
                <h3>No aligned figure or formula matches found matching the current filter. Ensure both PDF and Excel manifest are loaded.</h3>
            </div>
        `;
        return;
    }

    allMatched.forEach(item => {
        const isFigure = item.match_category === 'figure';
        const ex = item.excel_match;
        const confVal = (typeof item.confidence === 'number' && !isNaN(item.confidence)) ? item.confidence : (ex ? 0.95 : 0.85);
        const confPct = Math.round(confVal * 100);
        const isSelected = isFigure ? selectedFigureIds.has(item.figure_id) : selectedFormulaIds.has(item.formula_id);
        const isInjected = item.status_label === 'Injected';
        const effectiveAlt = (ex && ex.alt_text) || item.alt_text || (isFigure ? '' : item.actual_text) || '';
        const mcidText = item.mcids && item.mcids.length > 0 ? `MCID ${item.mcids.join(',')}` : (isFigure ? 'Structure Element' : (item.formula_type || '/Formula'));

        const card = document.createElement('div');
        card.className = `match-card ${isFigure ? '' : 'formula-match-card'} ${isSelected ? 'selected' : ''}`;
        if (isFigure) {
            card.setAttribute('data-fig-id', item.figure_id);
        } else {
            card.setAttribute('data-formula-id', item.formula_id);
        }

        const typeBadge = isFigure
            ? `<span class="match-type-badge figure-badge"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg> Figure Match</span>`
            : `<span class="match-type-badge formula-badge"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M4 19L9 5H20M12 12L15 19L21 9"/></svg> Formula Match</span>`;

        const selectToggleHandler = isFigure
            ? `window.toggleFigureSelection(${item.figure_id})`
            : `window.toggleFormulaSelection(${item.formula_id})`;

        const selectCheckboxHandler = isFigure
            ? `window.toggleFigureSelection(${item.figure_id}, this.checked)`
            : `window.toggleFormulaSelection(${item.formula_id}, this.checked)`;

        const dualZoomHandler = isFigure
            ? `window.expandFigureDual(${item.figure_id})`
            : (ex && ex.image_url ? `window.expandFormulaDual(${item.formula_id})` : `window.expandFormulaSingle(${item.formula_id})`);

        const inspectHandler = isFigure
            ? `window.openPdfModalById(${item.figure_id})`
            : `window.openFormulaModalById(${item.formula_id})`;

        const copyHandler = isFigure
            ? `window.copyFigureAlt(${item.figure_id})`
            : `window.copyFormulaAlt(${item.formula_id})`;

        const itemTitle = isFigure
            ? `PDF Figure #${item.figure_id}`
            : `PDF Formula #${item.formula_id}`;

        const dataAttr = isFigure
            ? `data-fig-id="${item.figure_id}"`
            : `data-formula-id="${item.formula_id}"`;

        // Right side (Excel) Image rendering
        let excelImgHtml = '';
        if (ex && ex.image_url) {
            excelImgHtml = `<img src="${ex.image_url}" alt="Excel match" loading="lazy">`;
        } else if (ex) {
            excelImgHtml = `
                <div class="excel-no-image-placeholder">
                    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="21" y1="21" x2="3" y2="3"/></svg>
                    <span>Text-Only Alt Record</span>
                </div>
            `;
        } else {
            excelImgHtml = `
                <div class="excel-no-image-placeholder">
                    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><path d="M9 15l2 2 4-4"></path></svg>
                    <span>Injected Custom Alt</span>
                </div>
            `;
        }

        const excelHeader = ex
            ? `<span class="match-side-title">Excel Row ${ex.row} (Sr. ${ex.sr_no || ex.row - 1})</span><span class="row-chip" style="position:static;">${escapeHtml(ex.filename || (isFigure ? 'Drawing' : 'Formula'))}</span>`
            : `<span class="match-side-title">Injected Alt Source</span><span class="row-chip" style="position:static;">Custom Injected</span>`;

        const excelAlt = ex
            ? (ex.alt_text || 'No ALT text in Excel')
            : (item.alt_text || 'Injected directly into StructTree');

        card.innerHTML = `
            <!-- PDF Item Side -->
            <div class="match-side" onclick="${dualZoomHandler}" style="cursor: pointer;" title="Click to open Side-by-Side Zoom">
                <div class="match-side-header">
                    <label class="match-select-label" onclick="event.stopPropagation();" title="Select ${itemTitle} for injection">
                        <input type="checkbox" class="match-card-checkbox custom-checkbox" ${dataAttr} ${isSelected ? 'checked' : ''} onchange="${selectCheckboxHandler}">
                        <span class="match-side-title">${itemTitle}</span>
                    </label>
                    <div style="display:flex; align-items:center; gap:6px; flex-shrink:0; white-space:nowrap;">
                        ${typeBadge}
                        <button class="btn-card-select-toggle ${isSelected ? 'selected' : ''}" onclick="event.stopPropagation(); ${selectToggleHandler}" title="Toggle selection for injection">
                            ${isSelected ? '✓ Selected' : '+ Select'}
                        </button>
                        <span class="page-chip" style="position:static; flex-shrink:0;">Page ${item.page_number}</span>
                    </div>
                </div>
                <div class="match-img-frame">
                    <img src="${item.image_url || ''}" alt="${itemTitle}" loading="lazy">
                </div>
                <div class="figure-alt-preview ${item.alt_text ? '' : 'empty'}">
                    <strong>${isInjected ? (isFigure ? '✓ Injected StructTree Alt:' : '✓ Injected Formula Alt:') : 'Current PDF Alt:'}</strong> ${escapeHtml(item.alt_text || 'None (Missing /Alt in PDF StructTree)')}
                </div>
            </div>

            <!-- Match Divider -->
            <div class="match-divider">
                <button class="btn btn-compare-expand btn-sm" onclick="event.stopPropagation(); ${dualZoomHandler}" title="Side-by-Side Zoom & Compare">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/></svg>
                    <span>Side-by-Side Zoom</span>
                </button>
            </div>

            <!-- Excel Image Side -->
            <div class="match-side" onclick="${dualZoomHandler}" style="cursor: pointer;" title="Click to open Side-by-Side Zoom">
                <div class="match-side-header">
                    ${excelHeader}
                </div>
                <div class="match-img-frame">
                    ${excelImgHtml}
                </div>
                <div class="excel-alt-box-card" style="min-height:54px;">
                    <strong>Authoritative Alt:</strong> ${escapeHtml(excelAlt)}
                </div>
            </div>
        `;
        matchGrid.appendChild(card);
    });
}

// ==========================================
// MODAL DIALOGS
// ==========================================
window.openPdfModalById = function (figId) {
    const fig = currentFigures.find(f => f.figure_id === figId);
    if (fig) openPdfModal(fig);
};

function openPdfModal(fig) {
    currentModalFigure = fig;
    currentModalFormula = null;
    currentModalExcelRecord = null;
    modalTitle.textContent = `PDF Figure ${fig.figure_id} Accessibility Inspector`;
    modalSubtitle.textContent = `Page ${fig.page_number} • Marked Content ID: ${fig.mcids && fig.mcids.length ? fig.mcids.join(', ') : 'None'}`;
    modalImage.src = fig.image_url || '';
    modalImage.style.cursor = 'zoom-in';
    modalImage.onclick = () => window.expandFigureSingle(fig.figure_id);
    modalDownloadLink.href = fig.image_url || '';
    modalDownloadLink.setAttribute('download', fig.crop_filename || `figure_${fig.figure_id}.png`);

    if (modalTypeGroup) modalTypeGroup.style.display = 'none';
    if (modalTypeLabel) modalTypeLabel.textContent = 'Accessibility Tag';
    if (modalTypeDisplay) modalTypeDisplay.textContent = '/S /Figure (Tagged Structural Element)';

    const ex = fig.excel_match;
    const effectiveAlt = (ex && ex.alt_text) || fig.alt_text || '';
    const isInjected = fig.status_label === 'Injected';

    if (modalAltLabel) {
        modalAltLabel.textContent = 'Authoritative /Alt Attribute (Editable for PDF Injection)';
        modalAltLabel.style.display = 'none';
    }
    if (modalAltTextarea) {
        modalAltTextarea.value = effectiveAlt;
        modalAltTextarea.readOnly = false;
        modalAltTextarea.style.display = 'block';
    }
    if (modalAltDisplay) {
        modalAltDisplay.style.display = 'none';
    }

    if (modalSaveAltBtn) {
        modalSaveAltBtn.style.display = 'inline-flex';
        modalSaveAltBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Save Alt Text`;
    }

    if (modalInjectBtn) {
        modalInjectBtn.style.display = 'inline-flex';
        modalInjectBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 4px; vertical-align: middle;"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><path d="M9 15l2 2 4-4"></path></svg>${isInjected ? 'Re-Inject into PDF' : 'Inject into PDF'}`;
    }

    if (modalBBoxGroup) modalBBoxGroup.style.display = 'none';
    if (modalPathGroup) modalPathGroup.style.display = 'none';
    if (modalMcidGroup) modalMcidGroup.style.display = 'none';
    if (modalMatchedExcelBox) modalMatchedExcelBox.style.display = 'none';

    figureModal.style.display = 'flex';
}

window.openFormulaModalById = function (formulaId) {
    const formula = currentFormulas.find(f => f.formula_id === formulaId);
    if (formula) openFormulaModal(formula);
};

function openFormulaModal(formula) {
    currentModalFigure = null;
    currentModalFormula = formula;
    currentModalExcelRecord = null;
    modalTitle.textContent = `PDF Formula #${formula.formula_id} Inspector`;
    modalSubtitle.textContent = `Page ${formula.page_number} • Tag: ${formula.formula_type || '/Formula'} • MCID: ${formula.mcids && formula.mcids.length ? formula.mcids.join(', ') : 'None'}`;
    modalImage.src = formula.image_url || '';
    modalImage.style.cursor = 'zoom-in';
    modalImage.onclick = () => window.expandFormulaSingle(formula.formula_id);
    modalDownloadLink.href = formula.image_url || '';
    modalDownloadLink.setAttribute('download', formula.crop_filename || `formula_${formula.formula_id}.png`);

    if (modalTypeGroup) modalTypeGroup.style.display = 'none';
    if (modalTypeLabel) modalTypeLabel.textContent = 'Accessibility Math Tag';
    if (modalTypeDisplay) modalTypeDisplay.textContent = `${formula.formula_type || '/Formula'} (Structural Element • ${formula.underlying_xobjects_count || 0} XObjects)`;

    const ex = formula.excel_match;
    const effectiveAlt = (ex && ex.alt_text) || formula.alt_text || formula.actual_text || '';
    const isInjected = formula.status_label === 'Injected';

    if (modalAltLabel) {
        modalAltLabel.textContent = 'Authoritative Formula /Alt Text (Editable for PDF Injection)';
        modalAltLabel.style.display = 'none';
    }
    if (modalAltTextarea) {
        modalAltTextarea.value = effectiveAlt;
        modalAltTextarea.readOnly = false;
        modalAltTextarea.style.display = 'block';
    }
    if (modalAltDisplay) {
        modalAltDisplay.style.display = 'none';
    }

    if (modalSaveAltBtn) {
        modalSaveAltBtn.style.display = 'inline-flex';
        modalSaveAltBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Save Alt Text`;
    }

    if (modalInjectBtn) {
        modalInjectBtn.style.display = 'inline-flex';
        modalInjectBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 4px; vertical-align: middle;"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><path d="M9 15l2 2 4-4"></path></svg>${isInjected ? 'Re-Inject into PDF' : 'Inject into PDF'}`;
    }

    if (modalBBoxGroup) modalBBoxGroup.style.display = 'none';
    if (modalPathGroup) modalPathGroup.style.display = 'none';
    if (modalMcidGroup) modalMcidGroup.style.display = 'none';
    if (modalMatchedExcelBox) modalMatchedExcelBox.style.display = 'none';

    figureModal.style.display = 'flex';
}

function openExcelModal(rec) {
    currentModalFigure = null;
    currentModalFormula = null;
    currentModalExcelRecord = rec;
    modalTitle.textContent = `Excel Drawing Row ${rec.row} (Sr. ${rec.sr_no})`;
    modalSubtitle.textContent = `Manifest File: ${rec.filename || 'Drawing'} • Authoritative Alt`;
    modalImage.src = rec.image_url || '';
    modalImage.style.cursor = 'zoom-in';
    modalImage.onclick = () => window.expandExcelRecord(rec.row);
    modalDownloadLink.href = rec.image_url || '';
    modalDownloadLink.setAttribute('download', rec.image_filename || `excel_row_${rec.row}.png`);

    if (modalTypeGroup) modalTypeGroup.style.display = 'none';
    if (modalTypeLabel) modalTypeLabel.textContent = 'Manifest Entry';
    if (modalTypeDisplay) modalTypeDisplay.textContent = `Excel Drawing Object • Row ${rec.row} • Sr. No ${rec.sr_no}`;

    if (modalAltLabel) {
        modalAltLabel.textContent = 'Authoritative Client ALT Text';
        modalAltLabel.style.display = 'none';
    }
    if (modalAltTextarea) {
        modalAltTextarea.value = rec.alt_text || 'No ALT text present in Excel record.';
        modalAltTextarea.readOnly = true;
        modalAltTextarea.style.display = 'block';
    }
    if (modalAltDisplay) {
        modalAltDisplay.style.display = 'none';
    }

    if (modalSaveAltBtn) {
        modalSaveAltBtn.style.display = 'none';
    }

    if (modalInjectBtn) {
        modalInjectBtn.style.display = 'none';
    }

    if (modalBBoxGroup) modalBBoxGroup.style.display = 'none';
    if (modalPathGroup) modalPathGroup.style.display = 'none';
    if (modalMcidGroup) modalMcidGroup.style.display = 'none';
    if (modalMatchedExcelBox) modalMatchedExcelBox.style.display = 'none';

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

// ==========================================
// PROGRESS UPLOAD & LOADING MANAGEMENT
// ==========================================
function uploadWithProgress(url, formData, options = {}) {
    const { onProgress, onUploadComplete, uploadId } = options;
    return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        currentUploadXHR = xhr;
        currentUploadId = uploadId;

        xhr.upload.addEventListener('progress', (e) => {
            if (e.lengthComputable && e.total > 0) {
                const percent = Math.min(100, Math.round((e.loaded / e.total) * 100));
                if (onProgress) {
                    onProgress(percent, e.loaded, e.total);
                }
            }
        });

        xhr.upload.addEventListener('load', () => {
            if (onUploadComplete) {
                onUploadComplete();
            }
        });

        xhr.addEventListener('load', () => {
            currentUploadXHR = null;
            currentUploadId = null;
            if (xhr.status >= 200 && xhr.status < 300) {
                try {
                    const data = JSON.parse(xhr.responseText);
                    resolve(data);
                } catch (err) {
                    reject(new Error('Invalid response from server'));
                }
            } else if (xhr.status === 499) {
                const abortError = new Error('Upload cancelled');
                abortError.name = 'AbortError';
                reject(abortError);
            } else {
                let errorMsg = 'Upload failed';
                try {
                    const err = JSON.parse(xhr.responseText);
                    errorMsg = err.detail || errorMsg;
                } catch (e) {
                    errorMsg = xhr.statusText || errorMsg;
                }
                reject(new Error(errorMsg));
            }
        });

        xhr.addEventListener('error', () => {
            currentUploadXHR = null;
            currentUploadId = null;
            reject(new Error('Network error during upload'));
        });

        xhr.addEventListener('abort', () => {
            currentUploadXHR = null;
            currentUploadId = null;
            const abortError = new Error('Upload cancelled');
            abortError.name = 'AbortError';
            reject(abortError);
        });

        xhr.open('POST', url);
        xhr.send(formData);
    });
}

function showLoading(title, status, options = {}) {
    const {
        isUpload = false,
        canCancel = false,
        onCancel = null,
        initialPercent = 0,
        initialPercentText = ''
    } = options;

    const loadingTitle = document.getElementById('loadingTitle');
    const loadingStatus = document.getElementById('loadingStatus');
    const loadingPercent = document.getElementById('loadingPercent');
    const progressBarFill = document.getElementById('progressBarFill');
    const loaderActions = document.getElementById('loaderActions');

    // Remember previous section to restore on cancel
    if (loadingSection.style.display === 'none') {
        previousVisibleSection = (resultsSection && resultsSection.style.display !== 'none') ? 'results' : 'upload';
    }

    if (loadingTitle) loadingTitle.textContent = title;
    if (loadingStatus) loadingStatus.textContent = status;

    if (loadingPercent) {
        if (initialPercentText) {
            loadingPercent.textContent = initialPercentText;
            loadingPercent.style.display = 'inline-block';
            loadingPercent.classList.remove('processing');
        } else if (isUpload) {
            loadingPercent.textContent = `${initialPercent}% uploaded`;
            loadingPercent.style.display = 'inline-block';
            loadingPercent.classList.remove('processing');
        } else {
            loadingPercent.style.display = 'none';
            loadingPercent.classList.remove('processing');
        }
    }

    if (progressBarFill) {
        if (isUpload) {
            progressBarFill.classList.remove('indeterminate');
            progressBarFill.style.width = `${initialPercent}%`;
        } else {
            progressBarFill.classList.add('indeterminate');
            progressBarFill.style.width = '40%';
        }
    }

    if (loaderActions) {
        loaderActions.style.display = canCancel ? 'flex' : 'none';
    }

    currentCancelHandler = onCancel;

    uploadSection.style.display = 'none';
    resultsSection.style.display = 'none';
    loadingSection.style.display = 'flex';
}

function updateLoadingProgress(percent, percentText, statusText) {
    const loadingPercent = document.getElementById('loadingPercent');
    const progressBarFill = document.getElementById('progressBarFill');
    const loadingStatus = document.getElementById('loadingStatus');

    if (progressBarFill) {
        progressBarFill.classList.remove('indeterminate');
        progressBarFill.style.width = `${Math.min(100, Math.max(0, percent))}%`;
    }

    if (loadingPercent) {
        loadingPercent.style.display = 'inline-block';
        if (percentText) {
            loadingPercent.textContent = percentText;
        } else {
            loadingPercent.textContent = `${percent}% uploaded`;
        }
        if (percent >= 100) {
            loadingPercent.classList.add('processing');
        } else {
            loadingPercent.classList.remove('processing');
        }
    }

    if (statusText && loadingStatus) {
        loadingStatus.textContent = statusText;
    }
}

function resetLoadingProgress() {
    const loadingPercent = document.getElementById('loadingPercent');
    const progressBarFill = document.getElementById('progressBarFill');
    if (loadingPercent) {
        loadingPercent.textContent = '0% uploaded';
        loadingPercent.classList.remove('processing');
        loadingPercent.style.display = 'none';
    }
    if (progressBarFill) {
        progressBarFill.classList.remove('indeterminate');
        progressBarFill.style.width = '0%';
    }
}

function hideLoading() {
    if (currentProgressPoller) {
        clearInterval(currentProgressPoller);
        currentProgressPoller = null;
    }
    loadingSection.style.display = 'none';
    resetLoadingProgress();
}

async function cancelCurrentUpload() {
    if (currentProgressPoller) {
        clearInterval(currentProgressPoller);
        currentProgressPoller = null;
    }
    const uploadIdToCancel = currentUploadId;
    const sid = currentSession ? currentSession.session_id : '';

    // 1. Abort browser upload request
    if (currentUploadXHR) {
        try {
            currentUploadXHR.abort();
        } catch (e) {}
        currentUploadXHR = null;
    }

    // 2. Stop backend processing if already started
    if (uploadIdToCancel || sid) {
        try {
            fetch('/api/cancel-excel-upload', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    upload_id: uploadIdToCancel,
                    session_id: sid
                })
            }).catch(() => {});
        } catch (e) {}
    }

    currentUploadId = null;

    // 3. Clear selected Excel and PDF file inputs
    if (excelFileInput) excelFileInput.value = '';
    if (pdfFileInput) pdfFileInput.value = '';

    // 4. Reset progress to 0%
    resetLoadingProgress();

    // 5. Return UI to the normal upload state
    hideLoading();
    if (previousVisibleSection === 'results' && currentSession) {
        resultsSection.style.display = 'block';
        uploadSection.style.display = 'none';
    } else {
        uploadSection.style.display = 'block';
        resultsSection.style.display = 'none';
    }
}

// ==========================================
// INTERACTIVE LIGHTBOX & ZOOM CONTROLLER
// ==========================================

function updateZoomDisplay() {
    if (zoomLevelDisplay) {
        zoomLevelDisplay.textContent = `${Math.round(lightboxZoom * 100)}%`;
    }
    applyTransform();
}

function applyTransform() {
    const singleContainer = document.querySelector('.lightbox-img-container');
    if (singleContainer) {
        singleContainer.style.transform = `translate(${lightboxPanX}px, ${lightboxPanY}px) scale(${lightboxZoom})`;
    }
    const dualGrid = document.querySelector('.lightbox-dual-grid');
    if (dualGrid) {
        const dualImgs = dualGrid.querySelectorAll('.lightbox-pane-content img');
        dualImgs.forEach(img => {
            img.style.transform = `scale(${lightboxZoom})`;
        });
    }
}

function zoomIn(step = 0.25) {
    if (lightboxZoom < 8.0) {
        lightboxZoom = Math.min(8.0, +(lightboxZoom + step).toFixed(2));
        updateZoomDisplay();
    }
}

function zoomOut(step = 0.25) {
    if (lightboxZoom > 0.3) {
        lightboxZoom = Math.max(0.3, +(lightboxZoom - step).toFixed(2));
        updateZoomDisplay();
    }
}

function resetZoom() {
    lightboxZoom = 1.0;
    lightboxPanX = 0;
    lightboxPanY = 0;
    updateZoomDisplay();
}

function toggleLightboxBg() {
    currentLightboxBgIndex = (currentLightboxBgIndex + 1) % lightboxBgModes.length;
    const mode = lightboxBgModes[currentLightboxBgIndex];
    if (lightboxStage) {
        lightboxBgModes.forEach(m => lightboxStage.classList.remove(m));
        lightboxStage.classList.add(mode);
    }
}

function setupLightboxPanning() {
    if (!lightboxStage) return;

    lightboxStage.addEventListener('mousedown', (e) => {
        // If clicking on a button or link, don't initiate pan
        if (e.target.closest('button') || e.target.closest('a')) return;
        isPanning = true;
        startPanX = e.clientX - lightboxPanX;
        startPanY = e.clientY - lightboxPanY;
        const container = document.querySelector('.lightbox-img-container');
        if (container) container.classList.add('panning');
    });

    window.addEventListener('mousemove', (e) => {
        if (!isPanning) return;
        lightboxPanX = e.clientX - startPanX;
        lightboxPanY = e.clientY - startPanY;
        applyTransform();
    });

    window.addEventListener('mouseup', () => {
        if (isPanning) {
            isPanning = false;
            const container = document.querySelector('.lightbox-img-container');
            if (container) container.classList.remove('panning');
        }
    });

    // Wheel zooming
    lightboxStage.addEventListener('wheel', (e) => {
        e.preventDefault();
        const delta = e.deltaY < 0 ? 0.2 : -0.2;
        if (delta > 0) zoomIn(0.2);
        else zoomOut(0.2);
    }, { passive: false });
}

function openLightboxSingle(opts) {
    const {
        title = 'Magnified Image View',
        subtitle = 'High-Resolution Visual Inspection',
        imageUrl = '',
        altText = '',
        itemId = null,
        itemType = 'generic',
        downloadName = 'image.png',
        canInject = false
    } = opts;

    currentLightboxData = opts;
    lightboxZoom = 1.0;
    lightboxPanX = 0;
    lightboxPanY = 0;

    if (lightboxTitle) lightboxTitle.textContent = title;
    if (lightboxSubtitle) lightboxSubtitle.textContent = subtitle;
    if (lightboxAltText) lightboxAltText.textContent = altText || 'No ALT text available.';
    if (lightboxDownloadLink) {
        lightboxDownloadLink.href = imageUrl;
        lightboxDownloadLink.setAttribute('download', downloadName);
    }
    if (lightboxInjectBtn) {
        lightboxInjectBtn.style.display = canInject && altText ? 'inline-flex' : 'none';
        lightboxInjectBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 4px; vertical-align: middle;"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><path d="M9 15l2 2 4-4"></path></svg>Inject Alt into PDF`;
    }

    if (lightboxStage) {
        lightboxStage.innerHTML = `
            <div class="lightbox-single-pane">
                <div class="lightbox-img-container">
                    <img src="${imageUrl}" alt="${escapeHtml(title)}" draggable="false">
                </div>
            </div>
        `;
    }

    updateZoomDisplay();
    if (imageLightboxModal) {
        imageLightboxModal.style.display = 'flex';
    }
}

function openLightboxDual(opts) {
    const {
        title = 'Side-by-Side Comparison',
        subtitle = 'Inspect Math Formula & Authoritative Drawing Side by Side',
        pdfImageUrl = '',
        excelImageUrl = '',
        pdfLabel = 'PDF Crop',
        excelLabel = 'Excel Drawing',
        altText = '',
        itemId = null,
        itemType = 'generic',
        downloadName = 'comparison.png',
        canInject = false
    } = opts;

    currentLightboxData = opts;
    lightboxZoom = 1.0;
    lightboxPanX = 0;
    lightboxPanY = 0;

    if (lightboxTitle) lightboxTitle.textContent = title;
    if (lightboxSubtitle) lightboxSubtitle.textContent = subtitle;
    if (lightboxAltText) lightboxAltText.textContent = altText || 'No ALT text available.';
    if (lightboxDownloadLink) {
        lightboxDownloadLink.href = pdfImageUrl || excelImageUrl;
        lightboxDownloadLink.setAttribute('download', downloadName);
    }
    if (lightboxInjectBtn) {
        lightboxInjectBtn.style.display = canInject && altText ? 'inline-flex' : 'none';
        lightboxInjectBtn.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 4px; vertical-align: middle;"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><path d="M9 15l2 2 4-4"></path></svg>Inject Alt into PDF`;
    }

    if (lightboxStage) {
        lightboxStage.innerHTML = `
            <div class="lightbox-dual-grid">
                <div class="lightbox-dual-pane">
                    <div class="lightbox-pane-header pdf-pane">
                        <span>${escapeHtml(pdfLabel)}</span>
                        <button class="btn btn-secondary btn-sm" onclick="window.openLightboxSingle({ title: '${escapeHtml(title)} (PDF Crop)', subtitle: '${escapeHtml(subtitle)}', imageUrl: '${pdfImageUrl}', altText: '${escapeHtml(altText)}', itemId: ${itemId}, itemType: '${itemType}', downloadName: '${downloadName}', canInject: ${canInject} })">
                            Solo Zoom
                        </button>
                    </div>
                    <div class="lightbox-pane-content">
                        <img src="${pdfImageUrl}" alt="PDF Crop" draggable="false">
                    </div>
                </div>
                <div class="lightbox-dual-pane">
                    <div class="lightbox-pane-header excel-pane">
                        <span>${escapeHtml(excelLabel)}</span>
                        <button class="btn btn-secondary btn-sm" onclick="window.openLightboxSingle({ title: '${escapeHtml(title)} (Excel Drawing)', subtitle: '${escapeHtml(subtitle)}', imageUrl: '${excelImageUrl}', altText: '${escapeHtml(altText)}', itemId: ${itemId}, itemType: '${itemType}', downloadName: '${downloadName}', canInject: ${canInject} })">
                            Solo Zoom
                        </button>
                    </div>
                    <div class="lightbox-pane-content">
                        <img src="${excelImageUrl}" alt="Excel Drawing" draggable="false">
                    </div>
                </div>
            </div>
        `;
    }

    updateZoomDisplay();
    if (imageLightboxModal) {
        imageLightboxModal.style.display = 'flex';
    }
}

function closeLightbox() {
    if (imageLightboxModal) {
        imageLightboxModal.style.display = 'none';
    }
    currentLightboxData = null;
    lightboxZoom = 1.0;
    lightboxPanX = 0;
    lightboxPanY = 0;
}

// Global window helpers for expanding from any card or modal
window.expandFormulaSingle = function (formulaId) {
    const formula = currentFormulas.find(f => f.formula_id === formulaId);
    if (!formula) return;
    const ex = formula.excel_match;
    const alt = (ex && ex.alt_text) || formula.alt_text || formula.actual_text || '';
    openLightboxSingle({
        title: `PDF Formula #${formula.formula_id} Crop`,
        subtitle: `Page ${formula.page_number} • Tag ${formula.formula_type || '/Formula'} • BBox: ${formula.bbox_width} × ${formula.bbox_height} pt`,
        imageUrl: formula.image_url,
        altText: alt,
        itemId: formulaId,
        itemType: 'formula',
        downloadName: formula.crop_filename || `formula_${formula.formula_id}.png`,
        canInject: Boolean(alt)
    });
};

window.expandFormulaExcel = function (formulaId) {
    const formula = currentFormulas.find(f => f.formula_id === formulaId);
    if (!formula || !formula.excel_match) return;
    const ex = formula.excel_match;
    openLightboxSingle({
        title: `Excel Manifest Math Drawing (Formula #${formulaId})`,
        subtitle: `Row ${ex.row} (Sr. ${ex.sr_no}) • File: ${ex.filename || 'Drawing'}`,
        imageUrl: ex.image_url,
        altText: ex.alt_text || '',
        itemId: formulaId,
        itemType: 'formula',
        downloadName: ex.image_filename || `excel_formula_${ex.row}.png`,
        canInject: Boolean(ex.alt_text)
    });
};

window.expandFormulaDual = function (formulaId) {
    const formula = currentFormulas.find(f => f.formula_id === formulaId);
    if (!formula || !formula.excel_match) return;
    const ex = formula.excel_match;
    const alt = ex.alt_text || formula.alt_text || formula.actual_text || '';
    openLightboxDual({
        title: `Formula #${formulaId} Side-by-Side Comparison`,
        subtitle: `Page ${formula.page_number} PDF Formula Crop vs. Excel Row ${ex.row} Manifest Drawing`,
        pdfImageUrl: formula.image_url,
        excelImageUrl: ex.image_url,
        pdfLabel: `PDF Formula #${formulaId} Crop (Page ${formula.page_number})`,
        excelLabel: `Excel Row ${ex.row} (Sr. ${ex.sr_no} • ${ex.filename || 'Drawing'})`,
        altText: alt,
        itemId: formulaId,
        itemType: 'formula',
        downloadName: formula.crop_filename || `formula_${formulaId}_compare.png`,
        canInject: Boolean(alt)
    });
};

window.expandFigureSingle = function (figId) {
    const fig = currentFigures.find(f => f.figure_id === figId);
    if (!fig) return;
    const ex = fig.excel_match;
    const alt = (ex && ex.alt_text) || fig.alt_text || '';
    openLightboxSingle({
        title: `PDF Figure ${fig.figure_id} Crop`,
        subtitle: `Page ${fig.page_number} • BBox: ${fig.bbox_width} × ${fig.bbox_height} pt • MCID: ${fig.mcids && fig.mcids.length ? fig.mcids.join(', ') : 'None'}`,
        imageUrl: fig.image_url,
        altText: alt,
        itemId: figId,
        itemType: 'figure',
        downloadName: fig.crop_filename || `figure_${fig.figure_id}.png`,
        canInject: Boolean(alt)
    });
};

window.expandFigureExcel = function (figId) {
    const fig = currentFigures.find(f => f.figure_id === figId);
    if (!fig || !fig.excel_match) return;
    const ex = fig.excel_match;
    openLightboxSingle({
        title: `Excel Manifest Image (Figure ${figId})`,
        subtitle: `Row ${ex.row} (Sr. ${ex.sr_no}) • File: ${ex.filename || 'Drawing'}`,
        imageUrl: ex.image_url,
        altText: ex.alt_text || '',
        itemId: figId,
        itemType: 'figure',
        downloadName: ex.image_filename || `excel_figure_${ex.row}.png`,
        canInject: Boolean(ex.alt_text)
    });
};

window.expandFigureDual = function (figId) {
    const fig = currentFigures.find(f => f.figure_id === figId);
    if (!fig || !fig.excel_match) return;
    const ex = fig.excel_match;
    const alt = ex.alt_text || fig.alt_text || '';
    openLightboxDual({
        title: `Figure ${figId} Side-by-Side Comparison`,
        subtitle: `Page ${fig.page_number} PDF Figure Crop vs. Excel Row ${ex.row} Manifest Drawing`,
        pdfImageUrl: fig.image_url,
        excelImageUrl: ex.image_url,
        pdfLabel: `PDF Figure ${figId} Crop (Page ${fig.page_number})`,
        excelLabel: `Excel Row ${ex.row} (Sr. ${ex.sr_no} • ${ex.filename || 'Drawing'})`,
        altText: alt,
        itemId: figId,
        itemType: 'figure',
        downloadName: fig.crop_filename || `figure_${figId}_compare.png`,
        canInject: Boolean(alt)
    });
};

window.expandExcelRecord = function (row) {
    const rec = currentExcelRecords.find(r => r.row === row);
    if (!rec || !rec.image_url) return;
    openLightboxSingle({
        title: `Excel Manifest Drawing (Row ${rec.row})`,
        subtitle: `Row ${rec.row} • Sr. ${rec.sr_no || 'N/A'} • File: ${rec.filename || 'Drawing'}`,
        imageUrl: rec.image_url,
        altText: rec.alt_text || '',
        itemId: rec.row,
        itemType: 'excel',
        downloadName: rec.image_filename || `excel_row_${rec.row}.png`,
        canInject: false
    });
};

window.openLightboxSingle = openLightboxSingle;
window.openLightboxDual = openLightboxDual;
window.closeLightbox = closeLightbox;

document.addEventListener('DOMContentLoaded', initEvents);

