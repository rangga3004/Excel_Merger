/**
 * Excel Merger Application
 * Menggabungkan file Excel multipart dari ZIP menjadi satu file
 */

// DOM Elements
const uploadZone = document.getElementById('uploadZone');
const fileInput = document.getElementById('fileInput');
const progressSection = document.getElementById('progressSection');
const progressTitle = document.getElementById('progressTitle');
const progressPercent = document.getElementById('progressPercent');
const progressFill = document.getElementById('progressFill');
const progressStatus = document.getElementById('progressStatus');
const fileListSection = document.getElementById('fileListSection');
const fileList = document.getElementById('fileList');
const resultSection = document.getElementById('resultSection');
const resultStats = document.getElementById('resultStats');
const downloadBtn = document.getElementById('downloadBtn');
const resetBtn = document.getElementById('resetBtn');
const errorSection = document.getElementById('errorSection');
const errorMessage = document.getElementById('errorMessage');
const errorResetBtn = document.getElementById('errorResetBtn');

// State
let mergedWorkbook = null;
let outputFileName = 'merged_excel.xlsx';

// Event Listeners
uploadZone.addEventListener('click', () => fileInput.click());
uploadZone.addEventListener('dragover', handleDragOver);
uploadZone.addEventListener('dragleave', handleDragLeave);
uploadZone.addEventListener('drop', handleDrop);
fileInput.addEventListener('change', handleFileSelect);
downloadBtn.addEventListener('click', downloadMergedFile);
resetBtn.addEventListener('click', resetApp);
errorResetBtn.addEventListener('click', resetApp);

/**
 * Handle drag over event
 */
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadZone.classList.add('drag-over');
}

/**
 * Handle drag leave event
 */
function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadZone.classList.remove('drag-over');
}

/**
 * Handle file drop event
 */
function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadZone.classList.remove('drag-over');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

/**
 * Handle file select from input
 */
function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

/**
 * Main function to process the ZIP file
 */
async function processFile(file) {
    // Validate file type
    if (!file.name.toLowerCase().endsWith('.zip')) {
        showError('File harus berformat .zip');
        return;
    }
    
    // Generate output filename based on input
    outputFileName = file.name.replace('.zip', '_merged.xlsx');
    
    // Show progress section, hide upload zone
    uploadZone.classList.add('hidden');
    showProgress();
    
    try {
        // Step 1: Read ZIP file
        updateProgress(10, 'Membaca file ZIP...');
        const zipData = await readFileAsArrayBuffer(file);
        
        // Step 2: Extract ZIP
        updateProgress(20, 'Mengekstrak file ZIP...');
        const zip = await JSZip.loadAsync(zipData);
        
        // Step 3: Find Excel files
        updateProgress(30, 'Mencari file Excel...');
        const excelFiles = [];
        
        zip.forEach((relativePath, zipEntry) => {
            if (!zipEntry.dir && relativePath.toLowerCase().endsWith('.xlsx')) {
                // Skip temporary/hidden Excel files
                if (!relativePath.startsWith('~$') && !relativePath.includes('/__MACOSX/')) {
                    excelFiles.push({
                        path: relativePath,
                        name: relativePath.split('/').pop(),
                        entry: zipEntry
                    });
                }
            }
        });
        
        if (excelFiles.length === 0) {
            throw new Error('Tidak ditemukan file Excel (.xlsx) dalam ZIP');
        }
        
        // Sort files naturally (Part1, Part2, ... Part10, Part11)
        excelFiles.sort((a, b) => naturalSort(a.name, b.name));
        
        // Step 4: Parse each Excel file
        updateProgress(40, `Memproses ${excelFiles.length} file Excel...`);
        
        const allData = [];
        let headerRow = null;
        let totalRows = 0;
        const fileStats = [];
        
        for (let i = 0; i < excelFiles.length; i++) {
            const excelFile = excelFiles[i];
            const progressPct = 40 + Math.round((i / excelFiles.length) * 40);
            updateProgress(progressPct, `Memproses: ${excelFile.name}`);
            
            // Read Excel file from ZIP
            const excelData = await excelFile.entry.async('arraybuffer');
            const workbook = XLSX.read(excelData, { type: 'array' });
            
            // Get first sheet
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            
            // Convert to JSON (array of arrays)
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            
            if (jsonData.length === 0) continue;
            
            let dataRows;
            if (headerRow === null) {
                // First file: take everything including header
                headerRow = jsonData[0];
                dataRows = jsonData;
            } else {
                // Subsequent files: skip header row
                dataRows = jsonData.slice(1);
            }
            
            const rowCount = dataRows.length - (headerRow === jsonData[0] ? 1 : 0);
            fileStats.push({
                name: excelFile.name,
                rows: rowCount
            });
            
            totalRows += rowCount;
            allData.push(...dataRows);
        }
        
        // Step 5: Display file list
        displayFileList(fileStats);
        
        // Step 6: Create merged workbook
        updateProgress(85, 'Membuat file Excel gabungan...');
        
        mergedWorkbook = XLSX.utils.book_new();
        const mergedSheet = XLSX.utils.aoa_to_sheet(allData);
        
        // Auto-size columns
        const colWidths = calculateColumnWidths(allData);
        mergedSheet['!cols'] = colWidths;
        
        XLSX.utils.book_append_sheet(mergedWorkbook, mergedSheet, 'Merged Data');
        
        // Step 7: Complete
        updateProgress(100, 'Selesai!');
        
        // Show result after a short delay
        setTimeout(() => {
            showResult(excelFiles.length, totalRows);
        }, 500);
        
    } catch (error) {
        console.error('Error processing file:', error);
        showError(error.message || 'Terjadi kesalahan saat memproses file');
    }
}

/**
 * Read file as ArrayBuffer
 */
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(new Error('Gagal membaca file'));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Natural sort for file names (handles numbers correctly)
 */
function naturalSort(a, b) {
    const ax = [], bx = [];
    
    a.replace(/(\d+)|(\D+)/g, (_, $1, $2) => { ax.push([$1 || Infinity, $2 || '']) });
    b.replace(/(\d+)|(\D+)/g, (_, $1, $2) => { bx.push([$1 || Infinity, $2 || '']) });
    
    while (ax.length && bx.length) {
        const an = ax.shift();
        const bn = bx.shift();
        const nn = (an[0] - bn[0]) || an[1].localeCompare(bn[1]);
        if (nn) return nn;
    }
    
    return ax.length - bx.length;
}

/**
 * Calculate optimal column widths
 */
function calculateColumnWidths(data) {
    if (!data || data.length === 0) return [];
    
    const colWidths = [];
    const maxWidth = 50;
    const minWidth = 10;
    
    // Get max columns
    const maxCols = Math.max(...data.map(row => (row ? row.length : 0)));
    
    for (let col = 0; col < maxCols; col++) {
        let maxLen = minWidth;
        
        // Sample first 100 rows for performance
        const sampleSize = Math.min(data.length, 100);
        for (let row = 0; row < sampleSize; row++) {
            if (data[row] && data[row][col] !== undefined) {
                const cellLen = String(data[row][col]).length;
                if (cellLen > maxLen) maxLen = cellLen;
            }
        }
        
        colWidths.push({ wch: Math.min(maxLen + 2, maxWidth) });
    }
    
    return colWidths;
}

/**
 * Display file list with row counts
 */
function displayFileList(fileStats) {
    fileList.innerHTML = fileStats.map(file => `
        <div class="file-item">
            <div class="file-item-icon">
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                    <path d="M3 2H10L13 5V14H3V2Z" stroke="currentColor" stroke-width="1.5"/>
                    <path d="M10 2V5H13" stroke="currentColor" stroke-width="1.5"/>
                </svg>
            </div>
            <span class="file-item-name">${file.name}</span>
            <span class="file-item-rows">${file.rows.toLocaleString()} rows</span>
        </div>
    `).join('');
    
    fileListSection.classList.add('active');
}

/**
 * Update progress UI
 */
function updateProgress(percent, status) {
    progressPercent.textContent = `${percent}%`;
    progressFill.style.width = `${percent}%`;
    progressStatus.textContent = status;
}

/**
 * Show progress section
 */
function showProgress() {
    progressSection.classList.add('active');
    resultSection.classList.remove('active');
    errorSection.classList.remove('active');
    updateProgress(0, 'Memulai...');
}

/**
 * Show result section
 */
function showResult(fileCount, rowCount) {
    progressSection.classList.remove('active');
    resultSection.classList.add('active');
    
    resultStats.innerHTML = `
        <div class="stat-item">
            <div class="stat-value">${fileCount}</div>
            <div class="stat-label">File Digabung</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">${rowCount.toLocaleString()}</div>
            <div class="stat-label">Total Baris</div>
        </div>
    `;
}

/**
 * Show error section
 */
function showError(message) {
    uploadZone.classList.add('hidden');
    progressSection.classList.remove('active');
    resultSection.classList.remove('active');
    fileListSection.classList.remove('active');
    errorSection.classList.add('active');
    errorMessage.textContent = message;
}

/**
 * Download merged file
 */
function downloadMergedFile() {
    if (!mergedWorkbook) {
        showError('Tidak ada file untuk didownload');
        return;
    }
    
    try {
        // Generate Excel file
        const wbout = XLSX.write(mergedWorkbook, { 
            bookType: 'xlsx', 
            type: 'array' 
        });
        
        // Create blob and download
        const blob = new Blob([wbout], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        saveAs(blob, outputFileName);
    } catch (error) {
        console.error('Error downloading file:', error);
        showError('Gagal mendownload file');
    }
}

/**
 * Reset application to initial state
 */
function resetApp() {
    // Reset state
    mergedWorkbook = null;
    outputFileName = 'merged_excel.xlsx';
    fileInput.value = '';
    
    // Reset UI
    uploadZone.classList.remove('hidden');
    progressSection.classList.remove('active');
    fileListSection.classList.remove('active');
    resultSection.classList.remove('active');
    errorSection.classList.remove('active');
    fileList.innerHTML = '';
    updateProgress(0, 'Menunggu...');
}

// Initialize
console.log('Excel Merger initialized');
