# Excel Checkpoint Generator API (Node.js)

Layanan REST API berbasis **Node.js (Express)** untuk menerima payload JSON data checkpoint / inspeksi dan menghasilkan file **Excel (.xlsx)** yang sangat rapi, ber-grouping/flat row, serta menyusun foto-foto checkpoint dengan ukuran seragam.

---

## 🚀 Cara Menjalankan Server Node.js

### 1. Mode Development (Auto-Reload via Nodemon)
Setiap kali Anda mengedit file kode, server akan otomatis **reload** tanpa perlu stop/start manual:
```bash
npm run dev
```

### 2. Mode Produksi Standard
```bash
npm start
```
Server Express akan berjalan pada `http://localhost:8000`.

---

## 🌟 Endpoints API yang Tersedia

1. **`POST /api/v1/generate-excel/toyota`** *(Format Laporan Toyota)*:
   - Menghasilkan laporan Excel khusus format **Toyota** (support nested JSON: `name`, `sections`, `items`, `checkPoint`, `hasilPenilaian`, `solution`, `hasilFoto`).
   - Termasuk kolom **Solusi / Tindakan Perbaikan** dan pewarnaan status badge Toyota (`O` -> Hijau, `X` -> Merah, `∆` -> Amber).
2. **`POST /api/v1/generate-excel/yamaha`** *(Format Laporan Yamaha)*:
   - Menghasilkan laporan Excel khusus format **Yamaha**.
3. **`POST /api/v1/generate-excel/:reportType`** *(Dinamis)*:
   - Endpoint modular untuk tipe merek lainnya.

---

## 📁 Struktur Proyek

```
excel_ai/
├── src/
│   ├── generators/
│   │   ├── index.js          # Registry & Router generator laporan
│   │   ├── toyotaGenerator.js# Generator Laporan Spesifik Toyota
│   │   └── yamahaGenerator.js# Generator Laporan Spesifik Yamaha
│   ├── imageHandler.js     # Meresize foto (120x90) & auto-prefix URL
│   ├── excelGenerator.js   # Main facade entry point
│   └── server.js           # REST API Express & Web UI Server
├── public/                 # Web UI (HTML, CSS, JS)
├── toyota_checkpoint_data.json # Contoh Data JSON Input Toyota
├── checkpoint_data.json     # Contoh Data JSON Input Yamaha
├── package.json            # Config dependensi & npm scripts (nodemon)
└── README.md                # Dokumentasi Penggunaan
```

---

## 📡 Contoh Request API Toyota (via `curl`)

```bash
curl -X 'POST' \
  'http://localhost:8000/api/v1/generate-excel/toyota' \
  -H 'Content-Type: application/json' \
  -d @toyota_checkpoint_data.json \
  --output Laporan_Inspeksi_Toyota.xlsx
```
