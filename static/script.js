// Frontend Controller for Myntra Auto Lister PRO

// DOM Elements
const body = document.body;
const themeCheckbox = document.getElementById('theme-checkbox');
const themeSun = document.getElementById('theme-sun');
const themeMoon = document.getElementById('theme-moon');

const itemPathInput = document.getElementById('item-path');
const contentPathInput = document.getElementById('content-path');
const templatePathInput = document.getElementById('template-path');
const outputDirInput = document.getElementById('output-dir');
const productImagesDirInput = document.getElementById('product-images-dir');
const useVisionCheckbox = document.getElementById('use-vision-checkbox');

const categorySearchInput = document.getElementById('category-search-input');
const categoryCombobox = document.getElementById('category-combobox');
const categoryDropdownList = document.getElementById('category-dropdown-list');

const pricingRuleSelect = document.getElementById('pricing-rule');
const discountGroup = document.getElementById('discount-group');
const discountSlider = document.getElementById('discount-slider');
const discountBubble = document.getElementById('discount-bubble');

const productsFilterText = document.getElementById('products-filter');

const btnGenerate = document.getElementById('btn-generate');
const btnDryRun = document.getElementById('btn-dry-run');

const progressWrapper = document.getElementById('progress-wrapper');
const progressStepText = document.getElementById('progress-step-text');
const progressPercentText = document.getElementById('progress-percent-text');
const progressBarFill = document.getElementById('progress-bar-fill');

const consoleTerminal = document.getElementById('console-terminal');
const globalStatusPill = document.getElementById('global-status-pill');
const globalStatusText = document.getElementById('global-status-text');

// State Variables
let isCloud = false;
let allCategories = [];
let currentLogFilter = 'ALL';
let allLogs = [];
let eventSource = null;
let statusInterval = null;

let lastRunId = '';
let lastOutputFilename = '';

// Temporary holder for uploaded image data during trainer session
let activeTempFrontPath = '';
let activeTempBackPath = '';

// Cloud UI Initialization
function enableCloudUI() {
    addConsoleLine('INFO', 'Application running in CLOUD MODE. Native file uploads active.');
    
    // Hide Output Save Folder Card
    const outputCard = document.getElementById('output-dir-card');
    if (outputCard) {
        outputCard.style.display = 'none';
    }
    
    // Customize Product Images Folder Card for ZIP uploads
    const imgTitle = document.getElementById('product-images-card-title');
    const imgDesc = document.getElementById('product-images-card-desc');
    const btnText = document.getElementById('btn-browse-product-images-text');
    
    if (imgTitle) imgTitle.innerHTML = 'Product Images ZIP <span class="optional-badge">Optional</span>';
    if (imgDesc) imgDesc.innerText = 'Upload a ZIP archive containing front/back product photos for batch AI extraction.';
    if (btnText) btnText.innerText = 'Select ZIP file';
    
    // Clear default local paths
    itemPathInput.value = '';
    contentPathInput.value = '';
    templatePathInput.value = '';
    outputDirInput.value = '';
    productImagesDirInput.value = '';
    
    // Update displays
    updateBadgeForPath('item-path');
    updateBadgeForPath('content-path');
    updateBadgeForPath('template-path');
    updateBadgeForPath('output-dir');
    updateBadgeForPath('product-images-dir');
}

// Asynchronous upload handler
window.uploadSelectedFile = async function(inputId, isZip = false) {
    const fileInput = document.getElementById(inputId + '-file');
    const loader = document.getElementById(inputId + '-loader');
    const inputField = document.getElementById(inputId);
    
    if (!fileInput || !fileInput.files.length) return;
    
    const file = fileInput.files[0];
    
    // Show loader and hide display badge
    if (loader) loader.style.display = 'flex';
    const displayField = document.getElementById(inputId + '-display');
    if (displayField) displayField.style.display = 'none';
    
    addConsoleLine('INFO', `Uploading ${file.name}...`);
    
    const formData = new FormData();
    formData.append('file', file);
    if (isZip) {
        formData.append('extract', 'true');
    }
    
    try {
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();
        
        if (response.status === 200) {
            // Store the path returned by server
            const finalPath = isZip ? data.extracted_dir : data.filepath;
            inputField.value = finalPath;
            
            // Show badge
            updateBadgeForPath(inputId);
            addConsoleLine('SUCCESS', `Successfully uploaded ${file.name} to server temporary storage.`);
        } else {
            throw new Error(data.error || 'Upload failed.');
        }
    } catch (e) {
        addConsoleLine('ERROR', `Upload of ${file.name} failed: ${e.message}`);
        alert(`Failed to upload file: ${e.message}`);
        fileInput.value = '';
    } finally {
        if (loader) loader.style.display = 'none';
    }
};

// Default values populated on start
const defaultPaths = {
    item: "G:\\My Drive\\Manan\\Automation\\Myntra\\ITEM DIRECTORY Main.xlsx",
    content: "",
    template: "G:\\My Drive\\Manan\\Automation\\Myntra\\Myntra-Sku-Template-2026-05-20.xlsx",
    output: "G:\\My Drive\\Manan\\Automation\\Myntra",
    product_images: ""
};

window.updateBadgeForPath = function(inputId) {
    const inputField = document.getElementById(inputId);
    const displayField = document.getElementById(inputId + '-display');
    if (inputField && displayField) {
        if (inputField.value) {
            displayField.style.display = 'flex';
            const span = displayField.querySelector('span');
            if (span) {
                const parts = inputField.value.split(/[\\/]/);
                span.innerText = parts[parts.length - 1] || inputField.value;
            }
        } else {
            displayField.style.display = 'none';
        }
    }
};

window.browsePathDirect = async function(inputId, type, title) {
    if (isCloud) {
        const fileInput = document.getElementById(inputId + '-file');
        if (fileInput) {
            fileInput.click();
        } else {
            addConsoleLine('ERROR', `File input for ${inputId} not found in Cloud Mode.`);
        }
        return;
    }
    const inputField = document.getElementById(inputId);
    const displayField = document.getElementById(inputId + '-display');
    const initialdir = inputField.value;
    
    addConsoleLine('INFO', `Triggering local ${type} browser for: ${title}...`);
    
    try {
        const response = await fetch('/api/browse', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ type, title, initialdir })
        });
        const data = await response.json();
        
        if (data.path) {
            inputField.value = data.path;
            if (displayField) {
                displayField.style.display = 'flex';
                const span = displayField.querySelector('span');
                if (span) {
                    const parts = data.path.split(/[\\/]/);
                    span.innerText = parts[parts.length - 1] || data.path;
                }
            }
            addConsoleLine('SUCCESS', `Path selected: "${data.path}"`);
        } else {
            addConsoleLine('INFO', `Browsing selection cancelled by user.`);
        }
    } catch (e) {
        addConsoleLine('ERROR', `Path selection failed: ${e.message}`);
    }
};

window.clearSelectedPath = function(inputId) {
    const inputField = document.getElementById(inputId);
    if (inputField) {
        inputField.value = '';
    }
    const displayField = document.getElementById(inputId + '-display');
    if (displayField) {
        displayField.style.display = 'none';
    }
    addConsoleLine('INFO', `Cleared path selection: ${inputId}`);
};

// 1. Initialize Application
document.addEventListener('DOMContentLoaded', async () => {
    // Initialize Lucide Icons
    lucide.createIcons();
    
    // Fetch configuration
    try {
        const res = await fetch('/api/config');
        const config = await res.json();
        isCloud = config.is_cloud;
    } catch (err) {
        console.error("Failed to load configuration:", err);
    }
    
    if (isCloud) {
        enableCloudUI();
    } else {
        // Set default paths
        itemPathInput.value = defaultPaths.item;
        contentPathInput.value = defaultPaths.content;
        templatePathInput.value = defaultPaths.template;
        outputDirInput.value = defaultPaths.output;
        productImagesDirInput.value = defaultPaths.product_images;
        
        // Update display badges on start if values exist
        updateBadgeForPath('item-path');
        updateBadgeForPath('content-path');
        updateBadgeForPath('template-path');
        updateBadgeForPath('output-dir');
        updateBadgeForPath('product-images-dir');
    }
    
    // Initialize interactive toggles
    setupThemeToggle();
    setupDropdownEvents();
    toggleDiscountWidget();
    setupDragAndDrop();
    
    // Load categories and defaults from API
    fetchCategories();
    
    // Load Gemini API Key status
    loadGeminiKeyStatus();
    
    // Establish connection to Server-Sent Event (SSE) logs
    connectLogStream();
    
    // Monitor status on startup in case server was running
    checkTaskStatus();
    
    // Restore last download card if available in localStorage
    const savedRunId = localStorage.getItem('myntra_last_run_id');
    const savedFilename = localStorage.getItem('myntra_last_filename');
    if (savedRunId && savedFilename) {
        setTimeout(() => {
            const globalStatus = globalStatusText ? globalStatusText.innerText.toLowerCase() : '';
            const isProgressActive = progressWrapper && progressWrapper.classList.contains('active');
            if (globalStatus !== 'running' && !isProgressActive) {
                showDownloadCard(savedRunId, savedFilename);
            }
        }, 1000);
    }
});

// 2. Theme Switching Logic (Light & Dark)
function setupThemeToggle() {
    themeCheckbox.addEventListener('change', () => {
        if (themeCheckbox.checked) {
            body.classList.remove('light-mode');
            body.classList.add('dark-mode');
            themeSun.classList.remove('active');
            themeMoon.classList.add('active');
        } else {
            body.classList.remove('dark-mode');
            body.classList.add('light-mode');
            themeMoon.classList.remove('active');
            themeSun.classList.add('active');
        }
    });
}

// 2b. Page Tab Navigation Switching Logic
window.switchTab = function(tabId) {
    // Hide all tab contents
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    
    // Deactivate all sidebar menu items
    document.querySelectorAll('.menu-item').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Show select tab content
    const targetContent = document.getElementById(tabId);
    if (targetContent) {
        targetContent.classList.add('active');
    }
    
    // Activate clicked button
    let btnId = '';
    if (tabId === 'welcome-page') btnId = 'btn-tab-welcome';
    else if (tabId === 'data-page') btnId = 'btn-tab-data';
    else if (tabId === 'settings-page') btnId = 'btn-tab-settings';
    else if (tabId === 'terminal-page') btnId = 'btn-tab-terminal';
    else if (tabId === 'vision-page') btnId = 'btn-tab-vision';
    
    const targetBtn = document.getElementById(btnId);
    if (targetBtn) {
        targetBtn.classList.add('active');
    }
    
    // Smoothly update header top-bar title text to match selected tab
    const titleText = document.getElementById('page-title-text');
    if (titleText) {
        if (tabId === 'welcome-page') titleText.innerText = 'Dashboard Overview';
        else if (tabId === 'data-page') titleText.innerText = 'Upload Center';
        else if (tabId === 'settings-page') titleText.innerText = 'Listing Settings';
        else if (tabId === 'terminal-page') titleText.innerText = 'Run Terminal';
        else if (tabId === 'vision-page') {
            titleText.innerText = 'AI Vision Database';
            loadTrainedLearnings();
        }
    }
};

window.clearFilterText = function() {
    if (productsFilterText) {
        productsFilterText.value = '';
    }
};

// 3. Searchable Custom Dropdown (Combobox) Logic
async function fetchCategories() {
    try {
        const response = await fetch('/api/categories');
        const data = await response.json();
        allCategories = data.categories || [];
        
        // Render initial category lists
        renderCategoryDropdown(allCategories);
        
        // Populate the vision category select in the Trainer tab
        const visionSelect = document.getElementById('vision-category-select');
        if (visionSelect) {
            visionSelect.innerHTML = '';
            allCategories.forEach(cat => {
                const opt = document.createElement('option');
                opt.value = cat;
                opt.innerText = cat;
                visionSelect.appendChild(opt);
            });
        }
        
        // Set default select category value to blank
        categorySearchInput.value = "";
    } catch (e) {
        addConsoleLine('ERROR', `Failed to load listing categories from server: ${e.message}`);
    }
}

function renderCategoryDropdown(items, filterText = '') {
    categoryDropdownList.innerHTML = '';
    
    const filtered = items.filter(item => 
        item.toLowerCase().includes(filterText.toLowerCase())
    );
    
    // Limit to top 15 matches to eliminate typing lag
    const displayedItems = filtered.slice(0, 15);
    
    if (displayedItems.length === 0) {
        const noResults = document.createElement('div');
        noResults.className = 'combobox-no-results';
        noResults.innerText = 'No matching categories';
        categoryDropdownList.appendChild(noResults);
        return;
    }
    
    displayedItems.forEach(cat => {
        const div = document.createElement('div');
        div.className = 'combobox-item';
        if (cat.toLowerCase() === categorySearchInput.value.toLowerCase()) {
            div.classList.add('selected');
        }
        
        // Highlight matching text to make it feel extremely modern and responsive!
        if (filterText) {
            const regex = new RegExp(`(${filterText.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')})`, 'gi');
            div.innerHTML = cat.replace(regex, `<span class="highlight-text">$1</span>`);
        } else {
            div.innerText = cat;
        }
        
        div.addEventListener('click', () => {
            categorySearchInput.value = cat;
            closeCombobox();
        });
        categoryDropdownList.appendChild(div);
    });
}

function setupDropdownEvents() {
    categorySearchInput.addEventListener('focus', () => {
        categoryCombobox.classList.add('open');
        renderCategoryDropdown(allCategories, ''); // Render ALL categories initially!
    });
    
    categorySearchInput.addEventListener('input', (e) => {
        renderCategoryDropdown(allCategories, e.target.value);
    });
    
    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (!categoryCombobox.contains(e.target)) {
            closeCombobox();
        }
    });
}

function closeCombobox() {
    categoryCombobox.classList.remove('open');
}

// 4. Form Pricing Widgets Logic
function toggleDiscountWidget() {
    const isDiscounted = pricingRuleSelect.value === 'Discounted';
    if (isDiscounted) {
        discountGroup.style.display = 'flex';
    } else {
        discountGroup.style.display = 'none';
    }
}

function updateSliderBubble(val) {
    discountBubble.innerText = `${val}%`;
}

// 5. HTML5 File Upload Selection Display
window.updateFileDisplay = function(inputId, displayId, fileNameId) {
    const fileInput = document.getElementById(inputId);
    const displayField = document.getElementById(displayId);
    const fileNameSpan = document.getElementById(fileNameId);
    
    if (fileInput && fileInput.files && fileInput.files.length > 0) {
        const file = fileInput.files[0];
        if (fileNameSpan) {
            fileNameSpan.innerText = file.name;
        }
        if (displayField) {
            displayField.style.display = 'flex';
        }
        addConsoleLine('SUCCESS', `Loaded file: ${file.name} (${formatBytes(file.size)})`);
    } else {
        if (displayField) {
            displayField.style.display = 'none';
        }
    }
};

function formatBytes(bytes, decimals = 2) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// 6. Real-time Log Streams using Server-Sent Events (SSE)
function connectLogStream() {
    if (eventSource) {
        eventSource.close();
    }
    
    eventSource = new EventSource('/api/logs');
    
    eventSource.onmessage = (event) => {
        try {
            const data = JSON.parse(event.data);

            // Silently ignore heartbeat events
            if (data.type === 'heartbeat') return;

            // Special download-trigger event: show card immediately from live stream
            if (data.level === '__DOWNLOAD__') {
                try {
                    const payload = JSON.parse(data.message);
                    showDownloadCard(payload.run_id, payload.filename);
                } catch(pe) {}
                return;
            }

            if (data.level && data.message) {
                addConsoleLine(data.level, data.message, data.time);
            }
        } catch (e) {
            // Ignore parse errors
        }
    };

    eventSource.onerror = () => {
        eventSource.close();
        setTimeout(connectLogStream, 4000);
    };
}

function addConsoleLine(level, message, timeStr = null) {
    const time = timeStr || new Date().toLocaleTimeString();
    const logObj = { time, level, message };
    allLogs.push(logObj);
    
    // Append to terminal container if it matches active filter
    if (currentLogFilter === 'ALL' || currentLogFilter === level) {
        appendLineToDOM(logObj);
    }
}

function appendLineToDOM(logObj) {
    const line = document.createElement('div');
    line.className = `console-line ${logObj.level.toLowerCase()}`;
    
    const timeSpan = document.createElement('span');
    timeSpan.className = 'c-time';
    timeSpan.innerText = `[${logObj.time}] `;
    
    const tagSpan = document.createElement('span');
    tagSpan.className = 'c-tag';
    tagSpan.innerText = `[${logObj.level}] `;
    
    const textSpan = document.createElement('span');
    textSpan.innerText = logObj.message;
    
    line.appendChild(timeSpan);
    line.appendChild(tagSpan);
    line.appendChild(textSpan);
    
    consoleTerminal.appendChild(line);
    
    // Cap at 200 items to eliminate scrolling and memory-repaint lag
    while (consoleTerminal.children.length > 200) {
        consoleTerminal.removeChild(consoleTerminal.firstChild);
    }
    
    consoleTerminal.scrollTop = consoleTerminal.scrollHeight;
}

function filterLogs(filterType) {
    currentLogFilter = filterType;
    
    // Update active button state
    document.querySelectorAll('.console-filter-btn').forEach(btn => {
        if (btn.innerText.toUpperCase() === filterType) {
            btn.classList.add('active');
        } else if (filterType === 'ALL' && btn.innerText === 'All') {
            btn.classList.add('active');
        } else if (filterType === 'WARNING' && btn.innerText === 'Warnings') {
            btn.classList.add('active');
        } else if (filterType === 'ERROR' && btn.innerText === 'Errors') {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
    
    // Re-render matching logs (capped to last 200 elements to prevent UI freeze)
    consoleTerminal.innerHTML = '';
    const filteredLogs = allLogs.filter(logObj => filterType === 'ALL' || logObj.level === filterType);
    const cappedLogs = filteredLogs.slice(-200);
    
    cappedLogs.forEach(logObj => {
        appendLineToDOM(logObj);
    });
}

function clearLogs() {
    allLogs = [];
    consoleTerminal.innerHTML = '';
    addConsoleLine('INFO', 'Log terminal cleared by user.');
}

// 7. Background Task Initiation and Tracking Logic
async function triggerAutomation(dryRun = false) {
    const useVision = useVisionCheckbox ? useVisionCheckbox.checked : false;
    
    // Gather form configurations
    const params = {
        item_path: itemPathInput.value.trim(),
        content_path: contentPathInput.value.trim(),
        template_path: templatePathInput.value.trim(),
        output_dir: outputDirInput.value.trim(),
        product_images_dir: productImagesDirInput.value.trim(),
        category: categorySearchInput.value.trim(),
        isp_rule: pricingRuleSelect.value,
        isp_discount: discountSlider.value,
        products_filter: productsFilterText.value.trim(),
        use_vision: useVision,
        dry_run: dryRun
    };
    
    // Input Validations
    if (!params.item_path || !params.template_path || (!isCloud && !params.output_dir)) {
        alert(isCloud ? 'Please upload Item Directory and SKU Template before executing!' : 'Please configure Item Directory, SKU Template, and Output Save Folder before executing!');
        return;
    }
    
    if (useVision && !params.product_images_dir) {
        addConsoleLine('WARNING', isCloud ? 'AI Vision is enabled but no Product Images ZIP is uploaded. Running listing without AI Vision.' : 'AI Vision is enabled but no Product Images Folder is selected. Running listing without AI Vision auto-fill.');
        params.use_vision = false;
        useVisionCheckbox.checked = false;
    }
    
    if (!params.category) {
        alert('Please select or search a Listing Category!');
        return;
    }
    
    // Hide download action card on new run
    const downloadCard = document.getElementById('download-action-card');
    if (downloadCard) {
        downloadCard.style.display = 'none';
    }
    // Clear last saved download card info from localStorage
    localStorage.removeItem('myntra_last_run_id');
    localStorage.removeItem('myntra_last_filename');
    
    // Toggle Button loadings
    btnGenerate.disabled = true;
    btnDryRun.disabled = true;
    
    progressWrapper.classList.add('active');
    updateProgressBar(0.0, "Launching local automation engine...");
    
    try {
        const response = await fetch('/api/run', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(params)
        });
        const data = await response.json();
        
        if (response.status !== 200) {
            throw new Error(data.error || 'Failed to trigger automation.');
        }
        
        addConsoleLine('INFO', `Automation task successfully submitted.`);
        
        // Start polling status
        if (statusInterval) clearInterval(statusInterval);
        statusInterval = setInterval(checkTaskStatus, 800);
        
    } catch (e) {
        addConsoleLine('ERROR', `Automation initiation aborted: ${e.message}`);
        btnGenerate.disabled = false;
        btnDryRun.disabled = false;
        progressWrapper.classList.remove('active');
    }
}

async function checkTaskStatus() {
    try {
        const response = await fetch('/api/status');
        const data = await response.json();
        
        updateStatusPill(data.status);
        
        if (data.status === 'running') {
            btnGenerate.disabled = true;
            btnDryRun.disabled = true;
            progressWrapper.classList.add('active');
            updateProgressBar(data.progress, data.current_step);
            
            // Auto-resume status polling on load/refresh if it is running but no interval exists
            if (!statusInterval) {
                statusInterval = setInterval(checkTaskStatus, 800);
            }
        } else if (data.status === 'success') {
            clearInterval(statusInterval);
            statusInterval = null;
            btnGenerate.disabled = false;
            btnDryRun.disabled = false;
            updateProgressBar(1.0, "Process completed successfully!");

            if (data.output_file) {
                showDownloadCard(data.run_id, data.output_file);
            } else {
                addConsoleLine('SUCCESS', `Dry run analysis finished successfully. Verified matching SKU rows.`);
            }
            
            setTimeout(() => {
                progressWrapper.classList.remove('active');
            }, 5000);
        } else if (data.status === 'failed') {
            clearInterval(statusInterval);
            statusInterval = null;
            btnGenerate.disabled = false;
            btnDryRun.disabled = false;
            updateProgressBar(1.0, `Error occurred: ${data.error}`);
            addConsoleLine('ERROR', `Task execution failed: ${data.error}`);
            setTimeout(() => {
                progressWrapper.classList.remove('active');
            }, 8000);
        } else {
            // Idle state
            clearInterval(statusInterval);
            statusInterval = null;
            btnGenerate.disabled = false;
            btnDryRun.disabled = false;
        }
    } catch (e) {
        console.error('Error fetching task status:', e);
    }
}

// 7b. Show download card and store run info
function showDownloadCard(runId, filename) {
    if (!runId || !filename) return;
    lastRunId = runId;
    lastOutputFilename = filename;
    
    // Store in localStorage so it persists across page reloads and server restarts
    localStorage.setItem('myntra_last_run_id', runId);
    localStorage.setItem('myntra_last_filename', filename);
    
    const downloadCard = document.getElementById('download-action-card');
    const downloadFileNameText = document.getElementById('download-file-name');
    if (downloadCard && downloadFileNameText) {
        downloadFileNameText.innerText = filename;
        downloadCard.style.display = 'block';
        downloadCard.scrollIntoView({ behavior: 'smooth' });
    }
    addConsoleLine('SUCCESS', `File ready for download: ${filename}`);
    lucide.createIcons();
}

// 7c. Download & Export Action Handlers
window.downloadPopulatedFile = function() {
    if (!lastRunId || !lastOutputFilename) {
        alert("No output file available to download!");
        return;
    }
    // Use the new persistent download endpoint (reads from MongoDB if file not on disk)
    const url = `/api/download-file?run_id=${encodeURIComponent(lastRunId)}&filename=${encodeURIComponent(lastOutputFilename)}`;
    window.location.href = url;
    addConsoleLine('INFO', `Initiated browser download for: ${lastOutputFilename}`);
};

window.exportAIBackup = function() {
    window.location.href = '/api/vision/export-db';
    addConsoleLine('INFO', 'Initiated backup export for trained Vision Database.');
};

function updateProgressBar(progressFloat, stepText) {
    const percent = Math.round(progressFloat * 100);
    progressBarFill.style.width = `${percent}%`;
    progressPercentText.innerText = `${percent}%`;
    progressStepText.innerText = stepText;
}

function updateStatusPill(status) {
    globalStatusPill.className = `status-pill ${status}`;
    if (status === 'idle') {
        globalStatusText.innerText = 'Ready';
    } else if (status === 'running') {
        globalStatusText.innerText = 'Running';
    } else if (status === 'success') {
        globalStatusText.innerText = 'Success';
    } else if (status === 'failed') {
        globalStatusText.innerText = 'Failed';
    }
}

// === AI Vision Trainer Features & Controller ===

window.toggleVisionSettings = function() {
    const enabled = useVisionCheckbox.checked;
    if (enabled) {
        addConsoleLine('INFO', 'AI Vision Auto-Fill enabled! Ensure you configure your Product Images Folder.');
    }
};

window.saveGeminiKey = async function() {
    const keyInput = document.getElementById('gemini-key-input');
    const key = keyInput.value.trim();
    
    if (!key) {
        alert('Please enter a valid Gemini API Key!');
        return;
    }
    
    try {
        const response = await fetch('/api/vision/save-key', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ key })
        });
        const data = await response.json();
        if (response.status === 200) {
            alert('Gemini API Key saved successfully!');
            keyInput.value = '';
            loadGeminiKeyStatus();
        } else {
            alert('Error: ' + data.error);
        }
    } catch (e) {
        alert('Failed to save API key: ' + e.message);
    }
};

async function loadGeminiKeyStatus() {
    const keyInput = document.getElementById('gemini-key-input');
    if (!keyInput) return;
    
    try {
        const response = await fetch('/api/vision/get-key');
        const data = await response.json();
        if (data.has_key) {
            keyInput.placeholder = `API Key Set (${data.masked_key})`;
        } else {
            keyInput.placeholder = 'Enter Google Gemini API Key...';
        }
    } catch (e) {
        console.error('Error fetching key status:', e);
    }
}

window.previewImage = function(input, previewId) {
    const previewContainer = document.getElementById(previewId);
    const file = input.files[0];
    
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            previewContainer.innerHTML = '';
            const img = document.createElement('img');
            img.src = e.target.result;
            img.className = 'preview-thumb';
            previewContainer.appendChild(img);
            previewContainer.parentElement.classList.add('has-image');
        };
        reader.readAsDataURL(file);
    }
};

function setupDragAndDrop() {
    const dropzones = document.querySelectorAll('.dropzone');
    dropzones.forEach(zone => {
        // Prevent default drag behaviors
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            zone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
            }, false);
        });
        
        // Highlights dropzone on hover
        ['dragenter', 'dragover'].forEach(eventName => {
            zone.addEventListener(eventName, () => zone.classList.add('highlight'), false);
        });
        ['dragleave', 'drop'].forEach(eventName => {
            zone.addEventListener(eventName, () => zone.classList.remove('highlight'), false);
        });
        
        // Handle dropped files
        zone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            const fileInput = zone.querySelector('input[type="file"]');
            
            if (files.length > 0) {
                fileInput.files = files;
                const previewId = zone.id === 'front-dropzone' ? 'front-preview' : 'back-preview';
                window.previewImage(fileInput, previewId);
            }
        }, false);
    });
}

window.classifyVisionProduct = async function() {
    const frontInput = document.getElementById('front-file-input');
    const backInput = document.getElementById('back-file-input');
    const category = document.getElementById('vision-category-select').value;
    const btn = document.getElementById('btn-classify-vision');
    
    if (!frontInput.files.length || !backInput.files.length) {
        alert('Please upload both Front and Back product images before running classification!');
        return;
    }
    
    // Toggle loading UI
    btn.disabled = true;
    const originalText = btn.innerHTML;
    btn.innerHTML = `<span class="spinner" style="display:inline-block; width:12px; height:12px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; margin-right:8px; vertical-align:middle;"></span><span>Analyzing Visuals (Calling Gemini)...</span>`;
    
    const formData = new FormData();
    formData.append('front_image', frontInput.files[0]);
    formData.append('back_image', backInput.files[0]);
    formData.append('category', category);
    
    if (templatePathInput && templatePathInput.value) {
        formData.append('template_path', templatePathInput.value);
    }
    
    try {
        const response = await fetch('/api/vision/analyze', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();
        
        if (response.status === 200) {
            // Save temp paths in state
            activeTempFrontPath = data.front_temp_path;
            activeTempBackPath = data.back_temp_path;
            
            renderPredictionResults(data.predictions, data.dropdowns);
        } else {
            alert('Classification Failed: ' + data.error);
        }
    } catch (e) {
        alert('Error communicating with vision server: ' + e.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalText;
    }
};

function renderPredictionResults(predictions, dropdowns) {
    const container = document.getElementById('vision-predicted-fields-container');
    const card = document.getElementById('vision-results-card');
    
    container.innerHTML = '';
    
    Object.keys(predictions).forEach(field => {
        const row = document.createElement('div');
        row.className = 'prediction-row';
        
        const label = document.createElement('span');
        label.className = 'prediction-label';
        label.innerText = field;
        
        const select = document.createElement('select');
        select.className = 'prediction-field-select';
        select.dataset.field = field;
        
        const opts = dropdowns[field] || [];
        opts.forEach(opt => {
            const el = document.createElement('option');
            el.value = opt;
            el.innerText = opt;
            if (opt === predictions[field]) {
                el.selected = true;
            }
            select.appendChild(el);
        });
        
        row.appendChild(label);
        row.appendChild(select);
        container.appendChild(row);
    });
    
    // Reset SKU input
    document.getElementById('vision-sku-input').value = '';
    
    // Show results panel
    card.style.display = 'block';
    
    // Re-initialize Lucide Icons
    lucide.createIcons();
    
    // Scroll to results
    card.scrollIntoView({ behavior: 'smooth' });
}

window.saveVisionLearning = async function() {
    const sku = document.getElementById('vision-sku-input').value.trim();
    const category = document.getElementById('vision-category-select').value;
    
    if (!sku) {
        alert('Please enter a valid Reference SKU Code (e.g. PGTOPW001956-MAROON) to train the model!');
        return;
    }
    
    // Gather predictions
    const labels = {};
    const selects = document.querySelectorAll('.prediction-field-select');
    selects.forEach(sel => {
        labels[sel.dataset.field] = sel.value;
    });
    
    const payload = {
        vendorArticleNumber: sku,
        category: category,
        front_temp_path: activeTempFrontPath,
        back_temp_path: activeTempBackPath,
        labels: labels
    };
    
    try {
        const response = await fetch('/api/vision/save-learning', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        const data = await response.json();
        
        if (response.status === 200) {
            alert('Model trained successfully on SKU ' + sku + '!');
            
            // Reset dropzones and prediction panel
            resetVisionTrainerForm();
            
            // Reload database gallery
            loadTrainedLearnings();
        } else {
            alert('Failed to save learning: ' + data.error);
        }
    } catch (e) {
        alert('Error saving trained weights: ' + e.message);
    }
};

function resetVisionTrainerForm() {
    document.getElementById('front-file-input').value = '';
    document.getElementById('back-file-input').value = '';
    
    const f_prev = document.getElementById('front-preview');
    const b_prev = document.getElementById('back-preview');
    
    f_prev.innerHTML = `<i data-lucide="image" class="dropzone-icon"></i><h5>Front Image</h5><p>Click or drag image here</p>`;
    b_prev.innerHTML = `<i data-lucide="image" class="dropzone-icon"></i><h5>Back Image</h5><p>Click or drag image here</p>`;
    
    f_prev.parentElement.classList.remove('has-image');
    b_prev.parentElement.classList.remove('has-image');
    
    document.getElementById('vision-results-card').style.display = 'none';
    
    activeTempFrontPath = '';
    activeTempBackPath = '';
    
    lucide.createIcons();
}

async function loadTrainedLearnings() {
    const gallery = document.getElementById('trained-gallery');
    if (!gallery) return;
    
    try {
        const response = await fetch('/api/vision/get-learnings');
        const data = await response.json();
        const items = data.learnings || [];
        
        gallery.innerHTML = '';
        
        if (items.length === 0) {
            gallery.innerHTML = `
                <div class="no-learnings-message">
                    <i data-lucide="info" style="width:32px; height:32px; color:var(--text-secondary); margin-bottom:8px;"></i>
                    <p>No trained examples found. Upload and save a classification to populate your customized learning database.</p>
                </div>
            `;
            lucide.createIcons();
            return;
        }
        
        items.forEach(item => {
            const card = document.createElement('div');
            card.className = 'trained-card';
            
            const imgContainer = document.createElement('div');
            imgContainer.className = 'trained-card-images';
            
            const f_img = document.createElement('img');
            f_img.src = item.front_image;
            f_img.className = 'trained-card-img';
            
            const b_img = document.createElement('img');
            b_img.src = item.back_image;
            b_img.className = 'trained-card-img';
            
            imgContainer.appendChild(f_img);
            imgContainer.appendChild(b_img);
            
            const info = document.createElement('div');
            info.className = 'trained-card-info';
            
            const h5 = document.createElement('h5');
            h5.innerText = item.vendorArticleNumber;
            
            const p = document.createElement('p');
            p.innerText = `Category: ${item.category}`;
            
            const tags = document.createElement('div');
            tags.className = 'trained-card-tags';
            
            Object.keys(item.labels).forEach(lbl => {
                const tag = document.createElement('span');
                tag.className = 'trained-tag';
                tag.innerText = `${lbl}: ${item.labels[lbl]}`;
                tags.appendChild(tag);
            });
            
            info.appendChild(h5);
            info.appendChild(p);
            info.appendChild(tags);
            
            const deleteBtn = document.createElement('button');
            deleteBtn.className = 'btn-delete-learning';
            deleteBtn.innerHTML = `<i data-lucide="trash-2" style="width: 16px; height: 16px;"></i>`;
            deleteBtn.title = 'Delete trained example';
            deleteBtn.addEventListener('click', () => deleteTrainedLearning(item.vendorArticleNumber));
            
            // Make card clickable to edit/correct predictions
            card.style.cursor = 'pointer';
            card.addEventListener('click', (e) => {
                if (e.target.closest('.btn-delete-learning')) return;
                editTrainedLearning(item);
            });
            
            card.appendChild(imgContainer);
            card.appendChild(info);
            card.appendChild(deleteBtn);
            gallery.appendChild(card);
        });
        
        lucide.createIcons();
    } catch (e) {
        console.error('Error loading database learnings:', e);
    }
}

async function deleteTrainedLearning(van) {
    if (!confirm(`Are you sure you want to delete the trained learning for SKU "${van}"?`)) {
        return;
    }
    
    try {
        const response = await fetch(`/api/vision/delete-learning/${van}`, {
            method: 'DELETE'
        });
        const data = await response.json();
        if (response.status === 200) {
            loadTrainedLearnings();
        } else {
            alert('Failed to delete learning: ' + data.error);
        }
    } catch (e) {
        alert('Error: ' + e.message);
    }
}

async function editTrainedLearning(item) {
    const container = document.getElementById('vision-predicted-fields-container');
    const cardPanel = document.getElementById('vision-results-card');
    
    // Show results panel with loading state
    cardPanel.style.display = 'block';
    container.innerHTML = `
        <div style="padding: 20px; text-align: center; color: var(--text-secondary); grid-column: span 2;">
            <span class="spinner" style="display:inline-block; width:16px; height:16px; border:2px solid var(--primary-color); border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; margin-right:8px; vertical-align:middle;"></span>
            Loading dropdown options for ${item.category}...
        </div>
    `;
    cardPanel.scrollIntoView({ behavior: 'smooth' });
    
    try {
        const url = `/api/vision/dropdowns?category=${encodeURIComponent(item.category)}`;
        const response = await fetch(url);
        const data = await response.json();
        
        if (response.status !== 200) {
            throw new Error(data.error || 'Failed to load dropdown choices.');
        }
        
        // Populate inputs
        activeTempFrontPath = item.front_image;
        activeTempBackPath = item.back_image;
        
        const f_prev = document.getElementById('front-preview');
        const b_prev = document.getElementById('back-preview');
        
        f_prev.innerHTML = `<img src="${item.front_image}" class="preview-thumb">`;
        b_prev.innerHTML = `<img src="${item.back_image}" class="preview-thumb">`;
        
        f_prev.parentElement.classList.add('has-image');
        b_prev.parentElement.classList.add('has-image');
        
        const categorySelect = document.getElementById('vision-category-select');
        if (categorySelect) {
            categorySelect.value = item.category;
        }
        
        document.getElementById('vision-sku-input').value = item.vendorArticleNumber;
        
        // Render prediction rows using data.dropdowns and setting selection to item.labels
        container.innerHTML = '';
        const fields = data.fields || [];
        const dropdowns = data.dropdowns || {};
        
        fields.forEach(field => {
            const row = document.createElement('div');
            row.className = 'prediction-row';
            
            const label = document.createElement('span');
            label.className = 'prediction-label';
            label.innerText = field;
            
            const select = document.createElement('select');
            select.className = 'prediction-field-select';
            select.dataset.field = field;
            
            const opts = dropdowns[field] || [];
            opts.forEach(opt => {
                const el = document.createElement('option');
                el.value = opt;
                el.innerText = opt;
                if (opt === item.labels[field]) {
                    el.selected = true;
                }
                select.appendChild(el);
            });
            
            row.appendChild(label);
            row.appendChild(select);
            container.appendChild(row);
        });
        
    } catch (e) {
        container.innerHTML = `
            <div style="padding: 20px; text-align: center; color: var(--danger-color); grid-column: span 2;">
                <i data-lucide="alert-circle" style="width:24px; height:24px; margin-bottom:8px; display:block; margin:0 auto 8px;"></i>
                Failed to load edit options: ${e.message}
            </div>
        `;
        lucide.createIcons();
    }
}
