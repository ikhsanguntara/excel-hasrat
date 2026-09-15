# Excel Inspection & Asset Counting AI Generator API (Node.js)

Layanan REST API & Web Dashboard modern berbasis **Node.js (Express), ExcelJS & Sharp** untuk mengonversi data JSON checkpoint inspeksi dan inventaris aset menjadi file **Excel (.xlsx)** berstandar profesional.

Dilengkapi dengan sistem pemrosesan foto paralel berkecepatan tinggi, in-memory caching ganda, deduplikasi request, serta penataan multi-foto (hingga 5 foto per baris) yang rapi dan proporsional.

---

## 🚀 Fitur Unggulan

- 🏎️ **Performa Super Cepat (< 500ms)**: Didukung oleh Sharp image processing paralel, deduplikasi request *in-flight*, serta *in-memory cache*.
- 🖼️ **Dukungan Multi-Foto (1–5 Foto per Row)**: Mendukung input array URL (`["url1", "url2", ...]`) maupun string dengan pemisah koma/baris baru, ditata horizontal otomatis dengan border dan latar seragam.
- 📑 **3 Format Laporan Terintegrasi**:
  1. **Form Asset Counting (PT Hasjrat Abadi)**: 12 kolom dengan pewarnaan presisi (Yellow `#FFFF00` & Soft Blue `#B8CCE4`), formatting angka mata uang (`#,##0`), dan highlight status barang.
  2. **Laporan Checkpoint Toyota**: Format *Section-Block Merged* dengan badge penilaian warna (`O` Hijau, `X` Merah, `∆` Amber) dan kolom solusi perbaikan.
  3. **Laporan Checkpoint Yamaha**: Format laporan inspeksi fasilitas dan showroom Yamaha standar.
- 🎯 **Filter Bertingkat (Cascading Filter)**: Filter fleksibel via query params atau Web UI berdasarkan Area, Section, dan Sub-bagian.
- 🖥️ **Web Dashboard Interaktif**: Drag-and-drop file JSON, toggle switch antar brand/format, preview metrik data, dan generator langsung dari browser.

---

## 💻 Cara Menjalankan Aplikasi

### 1. Instalasi Dependensi
```bash
npm install
```

### 2. Mode Development (Auto-Reload via Nodemon)
```bash
npm run dev
```

### 3. Mode Produksi Standard
```bash
npm start
```
Aplikasi berjalan pada: `http://localhost:8000`

---

## 📡 Daftar Lengkap API Endpoints

### 1. Generate Excel Endpoints (`POST`)

| Method | Endpoint | Format Laporan | Keterangan |
| :--- | :--- | :--- | :--- |
| `POST` | `/api/v1/generate-excel/asset-counting` | **Form Asset Counting** | Menghasilkan Excel Form Asset Counting 12 kolom |
| `POST` | `/api/v1/generate-excel/asset` | **Form Asset Counting** | *Alias singkat* untuk Form Asset Counting |
| `POST` | `/api/v1/generate-excel/toyota` | **Toyota Checkpoint** | Menghasilkan Excel Toyota (Section-Block Merged) |
| `POST` | `/api/v1/generate-excel/yamaha` | **Yamaha Checkpoint** | Menghasilkan Excel Yamaha Checkpoint standar |
| `POST` | `/api/v1/generate-excel/:reportType` | **Dinamis** | Menghasilkan Excel sesuai `:reportType` di parameter URL |
| `POST` | `/api/v1/generate-excel` | **Default (Yamaha)** | Fallback endpoint default |

#### 🔍 Query Parameters untuk Filter Data (Opsional):
- `?area=...` : Filter data berdasarkan nama Area (contoh: `?area=Area Showroom`)
- `?section=...` : Filter data berdasarkan Section (contoh: `?section=Parkir`)
- `?subdetail=...` : Filter data berdasarkan Sub-bagian

---

### 2. Endpoint Data Sampel & Dummy JSON (`GET`)

| Method | Endpoint | Deskripsi |
| :--- | :--- | :--- |
| `GET` | `/asset_counting_data.json` | Sampel 15 item Form Asset Counting lengkap |
| `GET` | `/dummy_asset_counting_15.json` | Dummy JSON Asset Counting 15 item (1–5 foto per baris) |
| `GET` | `/toyota_checkpoint_data.json` | Sampel data inspeksi checkpoint Toyota |
| `GET` | `/checkpoint_data.json` | Sampel data inspeksi checkpoint Yamaha |
| `GET` | `/dummy_checkpoint_15.json` | Dummy JSON Checkpoint 15 item (1–5 foto per baris) |

---

### 3. Web Dashboard Frontend (`GET`)

| Method | Endpoint | Deskripsi |
| :--- | :--- | :--- |
| `GET` | `/` | Web UI Dashboard interaktif untuk upload & konversi JSON ke Excel |

---

## 📋 Struktur Format JSON Input

### 1. Format Form Asset Counting (`asset-counting`)
```json
{
  "branch_name": "AMBON",
  "fiscal_year": 2026,
  "data": [
    {
      "no": 1,
      "nomor_asset_modul": "AST/10100/2021/001",
      "nomor_asset_scan": "01.01.02.01.0001",
      "nama_asset": "Laptop Dell Latitude 3420 Core i5",
      "tanggal_perolehan": "2021-03-15",
      "harga_perolehan": 14500000,
      "akumulasi_penyusutan": 7250000,
      "nbv": 7250000,
      "user_pengguna": "Ahmad Fauzi (IT Support)",
      "status_barang": "ADA",
      "img_path": [
        "https://hrms.hasjrat.co.id/horor/hororupload/section/27/39/1/1/sample1.jpg",
        "https://hrms.hasjrat.co.id/horor/hororupload/section/27/39/1/2/sample2.jpg"
      ],
      "keterangan": "Kondisi fisik mulus dan beroperasi normal"
    }
  ]
}
```

### 2. Format Toyota Checkpoint (`toyota`)
```json
[
  {
    "name": "Showroom Toyota",
    "sections": [
      {
        "no": 1,
        "name": "Monumen",
        "items": [
          {
            "checkPoint": "Kebersihan Monumen",
            "hasilPenilaian": "O",
            "solution": "Cukup dibersihkan rutin",
            "img_path": "https://hrms.hasjrat.co.id/horor/hororupload/section/27/39/1/1/sample.jpg"
          }
        ]
      }
    ]
  }
]
```

---

## 🛠️ Contoh Request API (cURL)

### 1. Generate Form Asset Counting:
```bash
curl -X POST http://localhost:8000/api/v1/generate-excel/asset-counting \
  -H "Content-Type: application/json" \
  -d @dummy_asset_counting_15.json \
  --output Form_Asset_Counting_AMBON.xlsx
```

### 2. Generate Checkpoint Toyota (dengan Filter Area):
```bash
curl -X POST "http://localhost:8000/api/v1/generate-excel/toyota?area=Showroom%20Toyota" \
  -H "Content-Type: application/json" \
  -d @toyota_checkpoint_data.json \
  --output Laporan_Toyota.xlsx
```

### 3. Generate Checkpoint Yamaha:
```bash
curl -X POST http://localhost:8000/api/v1/generate-excel/yamaha \
  -H "Content-Type: application/json" \
  -d @dummy_checkpoint_15.json \
  --output Laporan_Yamaha.xlsx
```

---

## 📁 Struktur Direktori Proyek

```
excel_ai/
├── src/
│   ├── generators/
│   │   ├── index.js                  # Registry & router generator laporan
│   │   ├── assetCountingGenerator.js # Generator Excel Form Asset Counting
│   │   ├── toyotaGenerator.js        # Generator Excel Checkpoint Toyota (Section-Block)
│   │   └── yamahaGenerator.js        # Generator Excel Checkpoint Yamaha
│   ├── imageHandler.js               # Pemrosesan gambar, thumbnail & in-memory cache
│   ├── excelGenerator.js             # Facade entry point
│   └── server.js                     # Express server & API routes
├── public/                           # Frontend Web UI (HTML, CSS, Vanilla JS)
│   ├── index.html
│   ├── style.css
│   └── app.js
├── asset_counting_data.json          # Contoh data Asset Counting (15 item)
├── dummy_asset_counting_15.json      # Dummy JSON Asset Counting (1-5 foto per row)
├── dummy_checkpoint_15.json          # Dummy JSON Checkpoint (1-5 foto per row)
├── toyota_checkpoint_data.json       # Contoh data Checkpoint Toyota
├── checkpoint_data.json              # Contoh data Checkpoint Yamaha
├── package.json
└── README.md                         # Dokumentasi Proyek
```

---

## 📄 Lisensi
Hak Cipta © 2026 PT Hasjrat Abadi. Dikembangkan untuk otomatisasi pelaporan operasional cabang.
