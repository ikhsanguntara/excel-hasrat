# Excel Checkpoint Generator API (Node.js)

Layanan REST API berbasis **Node.js (Express)** untuk menerima payload JSON data checkpoint / inspeksi dan menghasilkan file **Excel (.xlsx)** yang sangat rapi, ber-grouping, serta menyusun foto-foto checkpoint dengan ukuran seragam.

---

## 🌟 Fitur Utama

1. **REST API Endpoint (`POST /api/v1/generate-excel`)**:
   - Menerima payload JSON baik berupa **Array** `[{...}, {...}]` atau **Object** `{"data": [{...}]}`.
   - Mengembalikan stream file Excel `.xlsx` langsung sebagai attachment.
2. **Pengelompokan Visual (Grouping)**:
   - Data otomatis dikelompokkan secara hierarkis per **Area**, **Section**, dan **Sub-Section**.
   - Header berwarna kontras dengan styling profesional.
3. **Pengolahan & Alignment Foto (Sharp Image Resizing & Multi-Photo Grid)**:
   - Menggunakan **Sharp** untuk meresize foto secara presisi ke thumbnail **120x90 px** dengan border halus.
   - **Multi-Photo Row Arrangement**: Jika 1 checkpoint memiliki >1 foto (contoh: `"path1.jpg,path2.jpg"`), foto-foto disusun **berdampingan secara horizontal** per baris secara rapi di dalam 1 sel.
   - **Fallback Placeholder**: Jika foto tidak ditemukan atau path kosong, sistem secara otomatis menampilkan SVG placeholder (*Foto Tidak Ada*) tanpa membuat Excel crash.
4. **Desain Excel Profesional (ExcelJS)**:
   - Kartu Ringkasan Metadata Dokumen di bagian atas.
   - Custom font (Segoe UI), zebra row striping, auto column width, dan border halus.
   - Badge warna otomatis pada kolom Hasil ("Y" -> Hijau, "N" -> Merah).

---

## 📁 Struktur Proyek

```
excel_ai/
├── src/
│   ├── imageHandler.js     # Meresize foto (120x90), menyusun multi-foto side-by-side (Sharp)
│   ├── excelGenerator.js   # Generator Workbook Excel (ExcelJS) dengan Grouping & Styling
│   └── server.js           # REST API Express (POST /api/v1/generate-excel & GET /)
├── checkpoint_data.json     # Contoh Data JSON Input Checkpoint
├── test_generate_node.js   # Script Pengujian Lokal CLI Node.js
├── package.json            # Dependensi Node.js (express, exceljs, sharp, axios)
└── README.md                # Dokumentasi Penggunaan
```

---

## 🚀 Cara Menjalankan Server Node.js

### 1. Menjalankan Server
```bash
npm start
```
Server Express akan berjalan pada `http://localhost:8000`.

---

## 📡 Cara Menggunakan API (Contoh Request)

### Menggunakan `curl`:
```bash
curl -X 'POST' \
  'http://localhost:8000/api/v1/generate-excel' \
  -H 'Content-Type: application/json' \
  -d @checkpoint_data.json \
  --output Laporan_Inspeksi_Node.xlsx
```

---

## 🧪 Cara Menjalankan Uji Coba Lokal (CLI Test)

Untuk menggenerate file Excel langsung dari `checkpoint_data.json` tanpa server HTTP:
```bash
npm run test-excel
```
File hasil `laporan_checkpoint_node_sample.xlsx` akan otomatis dibuat di root proyek.
