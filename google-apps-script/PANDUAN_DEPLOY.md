# 📋 Panduan Deploy ke Google Apps Script

## Langkah 1: Buat Project Baru

1. Buka [Google Apps Script](https://script.google.com/)
2. Klik **"New Project"**
3. Ganti nama project menjadi "Excel Merger"

---

## Langkah 2: Setup File Code.gs

1. Di editor, sudah ada file `Code.gs` default
2. **Hapus semua isi** file tersebut
3. Buka file `Code.gs` dari folder ini
4. **Copy-paste** semua isinya ke editor

---

## Langkah 3: Buat File Index.html

1. Klik **"+"** di samping "Files"
2. Pilih **"HTML"**
3. Beri nama: `Index` (tanpa .html)
4. Buka file `Index.html` dari folder ini
5. **Copy-paste** semua isinya ke editor

---

## Langkah 4: Deploy sebagai Web App

1. Klik **"Deploy"** > **"New deployment"**
2. Klik ⚙️ (gear icon) > pilih **"Web app"**
3. Isi form:
   - **Description**: "Excel Merger v1.0"
   - **Execute as**: "Me"
   - **Who has access**: 
     - `"Anyone with [company domain]"` → Hanya karyawan
     - `"Anyone"` → Publik (siapa saja)
4. Klik **"Deploy"**
5. **Copy URL** yang muncul

---

## Langkah 5: Share ke Tim

1. Share URL ke tim via Slack, Email, atau bookmark
2. URL format: `https://script.google.com/macros/s/xxxxx/exec`

---

## 🔄 Update Aplikasi

Jika ada perubahan kode:
1. Edit file di Apps Script
2. Klik **"Deploy"** > **"Manage deployments"**
3. Klik ✏️ (edit) pada deployment
4. Pilih **"New version"**
5. Klik **"Deploy"**

---

## ⚠️ Troubleshooting

### Error: "Authorization required"
- User perlu klik "Review Permissions" dan allow akses

### Error: "Script function not found: doGet"
- Pastikan file `Code.gs` sudah ada function `doGet()`

### Tidak bisa load external library
- Pastikan koneksi internet stabil
- CDN libraries (JSZip, SheetJS) perlu diakses dari internet
