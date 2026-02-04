# Excel Tools (AHA Commerce) - Project Documentation

## Overview
**Excel Tools** is a web-based utility for processing and merging Excel files, deployed as a **Google Apps Script Web App**. It's designed for internal use at AHA Commerce.

**Live URL:** Stable production deployment with consistent URL across versions.

---

## Project Structure

```
📁 Program Pemersatu Part Excel/
├── 📁 google-apps-script/           # Main source code
│   ├── Index.html                   # Single-file web app (HTML + CSS + JS)
│   ├── Code.js                      # Apps Script backend (doGet)
│   ├── appsscript.json              # Manifest
│   ├── deploy.ps1                   # Deployment script
│   ├── .clasp.json                  # Clasp configuration
│   ├── PANDUAN_DEPLOY.md            # Deployment guide
│   └── PANDUAN_CLASP.md             # Clasp setup guide
└── 📁 .git/                         # Version control
```

---

## Core Features

### 1. **Dual Mode: Single File & ZIP Merge**
- **Single File Tab:** Upload one Excel file, filter columns, download
- **ZIP Merge Tab:** Upload ZIP containing multiple Excel files, merge into one

### 2. **Column Selection**
- Checkbox grid showing all columns
- Select All / Deselect All buttons
- Search/filter columns by name
- Real-time preview of selected columns

### 3. **Row Settings** *(Latest Feature)*
- **Header Row Selector:** Choose which row contains column headers (default: 1)
- **Data Start Row Selector:** Choose where data begins (default: 2)
- Enables skipping sub-headers or metadata rows
- Preview updates dynamically when changed

### 4. **Number Format Swap**
- Swap between ID format (1.234,56) and US format (1,234.56)
- Auto-detection of current format based on data patterns
- Visual indicator showing detected format

### 5. **Live Data Preview**
- Shows header row + 2 data rows
- Respects column selection (only shows selected columns)
- Respects row settings (uses configured header/data rows)
- Horizontal scrolling for wide tables

---

## Technical Architecture

### Frontend (Index.html)
- **Single-file architecture:** All HTML, CSS, and JavaScript in one file
- **Libraries (CDN):**
  - SheetJS (xlsx.full.min.js) - Excel parsing/writing
  - JSZip - ZIP file handling
  - FileSaver.js - Client-side file downloads
- **No framework:** Vanilla JS for simplicity and Apps Script compatibility

### Backend (Code.js)
```javascript
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Excel Tools - AHA Commerce')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```

### Key JavaScript Functions

| Function | Purpose |
|----------|---------|
| `processSingleFile()` | Handle single Excel upload |
| `processZipFile()` | Handle ZIP merge workflow |
| `showColumnSelector()` | Display column selection UI |
| `updatePreview()` | Refresh data preview (respects row/column settings) |
| `downloadMergedFile()` | Create and download filtered Excel |
| `swapNumberFormat()` | Toggle number format (ID ↔ US) |
| `handleRowSettingsChange()` | Handle header/data row changes |
| `refreshColumnHeaders()` | Update column grid when header row changes |
| `getSelectedColumns()` | Get array of selected column indices |
| `resetApp()` | Reset all state to initial |

### Key State Variables

```javascript
let allMergedData = [];       // All raw data from file(s)
let headerRowIndex = 0;       // 0-based index of header row
let dataStartRowIndex = 1;    // 0-based index where data starts
let formatSwapped = false;    // Track number format swap state
```

---

## Deployment Workflow

### Prerequisites
- Node.js installed
- Clasp CLI: `npm install -g @google/clasp`
- Logged in: `clasp login`

### Deploy Steps
```powershell
cd google-apps-script
./deploy.ps1
```

The script:
1. Pushes code via `clasp push`
2. Creates new deployment via `clasp deploy`
3. Uses stable deployment ID for consistent URL

### Version History (Recent)
| Version | Changes |
|---------|---------|
| #21 | Row selector feature (header row + data start row) |
| #20 | Fixed checkbox label click toggling |
| #19 | Dynamic preview based on column selection |
| #18 | Preview shows all columns with horizontal scroll |
| #17 | Added horizontal scroll to preview |
| #16 | Fixed orphaned session persistence code |
| #15 | Fixed upload bug (indentation + orphaned code) |

---

## Branding
- **Colors:** AHA Commerce blue palette (#2563EB, #1E40AF, etc.)
- **Font:** Manrope (Google Fonts)
- **Style:** Modern card-based UI with gradients and shadows

---

## Common Issues & Solutions

### Upload Not Working
**Symptoms:** Click upload button, file explorer doesn't open
**Cause:** JavaScript syntax error preventing event handlers from registering
**Solution:** Check for:
- Incorrect indentation in functions
- Orphaned code blocks (leftover from refactoring)
- Unclosed braces

### Preview Overflow
**Symptoms:** Preview table extends beyond container
**Solution:** Add `overflow-x: auto` to `.format-preview` CSS

### Double Toggle on Checkbox
**Symptoms:** Clicking label toggles checkbox twice
**Cause:** HTML `<label for="">` already toggles, plus JS handler also toggles
**Solution:** Exclude LABEL from manual toggle in click handler

---

## Future Enhancement Ideas
- Export to multiple formats (CSV, XLSX, etc.)
- Save column presets for reuse
- Dark mode toggle
- Drag-and-drop column reordering
- Multi-sheet support within single file

---

## Git Repository
**Remote:** GitHub (rangga3004/Excel_Merger)
**Branch:** main

---

*Last Updated: 2026-02-04*
*Current Version: #21*
