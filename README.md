# Excel Checkpoint Generator API (Node.js)

Layanan REST API berbasis **Node.js (Express)** untuk menerima payload JSON data checkpoint / inspeksi dan menghasilkan file **Excel (.xlsx)** yang sangat rapi, ber-grouping, serta menyusun foto-foto checkpoint dengan ukuran seragam.

---

## 🌟 Fitur & Arsitektur Generator Modular

1. **Endpoint Spesifik YAMAHA (`POST /api/v1/generate-excel/yamaha`)**:
   - Menghasilkan laporan Excel khusus format Yamaha.
2. **Endpoint Dinamis (`POST /api/v1/generate-excel/:reportType`)**:
   - Jika nanti ada format laporan baru (contoh: `honda`, `toyota`, dll.), cukup tambahkan file generator di folder `src/generators/` dan daftarkan di `src/generators/index.js`.
3. **Endpoint Default (`POST /api/v1/generate-excel`)**:
   - Tetap tersedia sebagai fallback alias default.

---

## 📁 Struktur Proyek Modular

```
excel_ai/
├── src/
│   ├── generators/
│   │   ├── index.js          # Registry & Router untuk berbagai tipe laporan
│   │   └── yamahaGenerator.js# Generator Laporan Spesifik Yamaha
│   ├── imageHandler.js     # Meresize foto (120x90) & auto-prefix https://hrms.hasjrat.co.id/horor/
│   ├── excelGenerator.js   # Facade Generator Utama
│   └── server.js           # REST API Express & Web UI Server
├── public/                 # Web UI (HTML, CSS, JS)
├── checkpoint_data.json     # Contoh Data JSON Input Checkpoint
├── test_generate_node.js   # Script Pengujian CLI Node.js
├── package.json            # Dependensi Node.js
└── README.md                # Dokumentasi Penggunaan
```

---

## 🚀 Cara Menjalankan Server Node.js

```bash
npm start
```
Server Express akan berjalan pada `http://localhost:8000`.

---

## 📡 Cara Menggunakan API Yamaha

### Menggunakan `curl`:
```bash
curl -X 'POST' \
  'http://localhost:8000/api/v1/generate-excel/yamaha' \
  -H 'Content-Type: application/json' \
  -d @checkpoint_data.json \
  --output Laporan_Inspeksi_Yamaha.xlsx
```
