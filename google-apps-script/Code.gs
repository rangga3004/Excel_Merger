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
