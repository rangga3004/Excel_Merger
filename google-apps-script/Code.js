/**
 * Excel Merger Web App - Google Apps Script
 * 
 * CARA DEPLOY:
 * 1. Buka https://script.google.com/
 * 2. Buat project baru
 * 3. Copy isi file ini ke Code.gs
 * 4. Buat file baru "Index.html" dan copy isi Index.html
 * 5. Click Deploy > New Deployment
 * 6. Pilih "Web app"
 * 7. Execute as: Me
 * 8. Who has access: "Anyone with [company] account" atau "Anyone"
 * 9. Click Deploy
 * 10. Copy URL dan share ke tim
 */

/**
 * Serve the main HTML page
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Excel Merger - Gabungkan File Excel dari ZIP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================
// GOOGLE SHEETS EXPORT FUNCTIONS
// ============================================

/**
 * Create a new Google Spreadsheet with the provided data
 * @param {Array<Array>} data - 2D array of data
 * @param {string} filename - Name for the new spreadsheet
 * @returns {Object} - {success: boolean, url?: string, error?: string}
 */
function createNewSpreadsheet(data, filename) {
  try {
    const spreadsheet = SpreadsheetApp.create(filename || 'Excel Merge Export');
    const sheet = spreadsheet.getActiveSheet();
    
    if (data && data.length > 0) {
      const numRows = data.length;
      const numCols = data[0].length;
      sheet.getRange(1, 1, numRows, numCols).setValues(data);
    }
    
    return {
      success: true,
      url: spreadsheet.getUrl(),
      id: spreadsheet.getId()
    };
  } catch (error) {
    return {
      success: false,
      error: 'Gagal membuat spreadsheet: ' + error.message
    };
  }
}

/**
 * Get list of sheet names from a Google Spreadsheet URL
 * @param {string} url - Google Spreadsheet URL
 * @returns {Object} - {success: boolean, sheets?: Array<string>, error?: string}
 */
function getSpreadsheetSheets(url) {
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(url);
    const sheets = spreadsheet.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    
    return {
      success: true,
      sheets: sheetNames,
      spreadsheetName: spreadsheet.getName()
    };
  } catch (error) {
    // Check for common errors
    const errorMsg = error.message.toLowerCase();
    if (errorMsg.includes('access') || errorMsg.includes('permission') || errorMsg.includes('denied')) {
      return {
        success: false,
        error: 'Anda tidak memiliki akses ke spreadsheet ini. Pastikan file di-share dengan akun Anda sebagai Editor.'
      };
    }
    if (errorMsg.includes('not found') || errorMsg.includes('invalid')) {
      return {
        success: false,
        error: 'URL tidak valid atau spreadsheet tidak ditemukan.'
      };
    }
    return {
      success: false,
      error: 'Gagal mengakses spreadsheet: ' + error.message
    };
  }
}

/**
 * Export data to an existing Google Spreadsheet
 * @param {string} url - Google Spreadsheet URL
 * @param {string} sheetName - Target sheet name
 * @param {number} startRow - Starting row (1-indexed)
 * @param {number} startCol - Starting column (1-indexed, A=1, B=2, etc.)
 * @param {Array<Array>} data - 2D array of data
 * @returns {Object} - {success: boolean, error?: string}
 */
function exportToExistingSheet(url, sheetName, startRow, startCol, data) {
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(url);
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      return {
        success: false,
        error: 'Sheet "' + sheetName + '" tidak ditemukan.'
      };
    }
    
    if (!data || data.length === 0) {
      return {
        success: false,
        error: 'Tidak ada data untuk diekspor.'
      };
    }
    
    const numRows = data.length;
    const numCols = data[0].length;
    const row = startRow || 1;
    const col = startCol || 1;
    
    // Clear existing data in the target range (for overwrite)
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow >= row && lastCol >= col) {
      const clearRows = Math.max(lastRow - row + 1, numRows);
      const clearCols = Math.max(lastCol - col + 1, numCols);
      sheet.getRange(row, col, clearRows, clearCols).clearContent();
    }
    
    // Write new data starting at specified cell
    sheet.getRange(row, col, numRows, numCols).setValues(data);
    
    return {
      success: true,
      url: spreadsheet.getUrl(),
      rowsWritten: numRows
    };
  } catch (error) {
    const errorMsg = error.message.toLowerCase();
    if (errorMsg.includes('access') || errorMsg.includes('permission') || errorMsg.includes('denied')) {
      return {
        success: false,
        error: 'Anda tidak memiliki akses Editor ke spreadsheet ini.'
      };
    }
    return {
      success: false,
      error: 'Gagal export ke spreadsheet: ' + error.message
    };
  }
}
