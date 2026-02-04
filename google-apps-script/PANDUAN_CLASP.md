# Setup Clasp untuk Google Apps Script

## Instalasi Global

Jalankan perintah ini untuk menginstall clasp secara global:

```bash
npm install -g @google/clasp
```

## Login ke Google Account

```bash
clasp login
```

Ini akan membuka browser untuk autentikasi dengan akun Google Anda.

## Mendapatkan Script ID

1. Buka project Apps Script Anda di: https://script.google.com
2. Klik **Project Settings** (ikon gear) di sidebar kiri
3. Copy **Script ID** yang ditampilkan

## Konfigurasi Script ID

Edit file `.clasp.json` dan ganti `YOUR_SCRIPT_ID_HERE` dengan Script ID yang sudah dicopy:

```json
{
  "scriptId": "1ABC...xyz",
  "rootDir": "."
}
```

## Perintah Clasp yang Berguna

### Push ke Apps Script
```bash
cd google-apps-script
clasp push
```

### Pull dari Apps Script
```bash
cd google-apps-script
clasp pull
```

### Buka di Browser
```bash
clasp open
```

### Deploy sebagai Web App (URL Baru)
```bash
clasp deploy --description "Versi terbaru"
```
> **Note:** Perintah ini akan membuat **URL BARU**. Gunakan hanya jika Anda ingin membuat instance baru.

#### Update Deployment (URL Tetap)
**Deployment ID Anda:** `AKfycbxLMEU6JFZPWLEmp7KM5x9zXnbaaYUc1A1rPK_wStoJ3_L1PfFa0NQSNRW1hH_-rce3`

Untuk mempermudah, saya sudah buatkan script. Cukup jalankan:
```bash
./deploy.ps1
```
atau double click `deploy.bat`

Jika ingin manual:
```bash
clasp deploy --deploymentId AKfycbxLMEU6JFZPWLEmp7KM5x9zXnbaaYUc1A1rPK_wStoJ3_L1PfFa0NQSNRW1hH_-rce3 --description "Update fitur baru"
```
Dengan cara ini, URL web app Anda **tidak akan berubah**.
```bash
clasp deployments
```

### Watch Mode (Auto Push saat File Berubah)
```bash
clasp push --watch
```

## Struktur File

```
google-apps-script/
├── .clasp.json        # Konfigurasi clasp (Script ID)
├── appsscript.json    # Manifest Apps Script
├── Code.gs            # Server-side script
├── Index.html         # Client-side UI
└── PANDUAN_CLASP.md   # Panduan ini
```

## Catatan Penting

- **Script ID** adalah private, jangan share ke public
- File `.clasprc.json` (credentials) sudah di-ignore oleh git
- Setiap kali push, file akan di-overwrite di Apps Script
- Gunakan `clasp pull` untuk mengambil perubahan dari Apps Script editor online
