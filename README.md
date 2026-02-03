# Excel Merger - AHA Commerce

Aplikasi web untuk menggabungkan file Excel multipart dari ZIP menjadi satu file.

## Fitur

- ✅ Upload ZIP via drag & drop
- ✅ Auto-extract file `.xlsx` dari ZIP
- ✅ Natural sorting (Part1, Part2, ..., Part10)
- ✅ Merge dengan single header
- ✅ Progress indicator
- ✅ Download hasil merge

## Cara Pakai

### Local (Browser)
1. Buka `index.html` di browser
2. Drag & drop file ZIP
3. Download hasil merge

### Google Apps Script
1. Buka [script.google.com](https://script.google.com)
2. Buat project baru
3. Copy `google-apps-script/Code.gs` dan `Index.html`
4. Deploy sebagai Web App
5. Share URL ke tim

Lihat `google-apps-script/PANDUAN_DEPLOY.md` untuk panduan lengkap.

## Tech Stack

- **JSZip** - Extract ZIP files
- **SheetJS** - Parse & create Excel
- **FileSaver.js** - Download files

## Struktur Project

```
├── index.html          # Local version
├── style.css           # Styling
├── app.js              # Core logic
└── google-apps-script/
    ├── Code.gs         # Apps Script server
    ├── Index.html      # All-in-one HTML
    └── PANDUAN_DEPLOY.md
```

## Branding

UI menggunakan AHA Commerce brand guidelines:
- Font: Manrope
- Primary: #2563EB (Blue)
- Cards: #FFFBF0 (Cream)
