const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

async function createMappingDoc() {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Excel AI Generator';
    workbook.created = new Date();

    const FONT_FAMILY = 'Segoe UI';

    const headerFillNavy = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    const headerFillBlue = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
    const fillZebra = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
    const fillHeaderYellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    const fillHeaderSoftBlue = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB8CCE4' } };
    const fillCellYellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

    const borderThin = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };

    const borderDark = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
    };

    // ==========================================
    // SHEET 1: KAMUS DATA & MAPPING JSON -> EXCEL
    // ==========================================
    const wsMapping = workbook.addWorksheet('1. Kamus Data & Mapping SA', { views: [{ showGridLines: true }] });

    // Judul Dokumen
    wsMapping.mergeCells('A1:G1');
    const titleCell = wsMapping.getCell('A1');
    titleCell.value = 'SPESIFIKASI MAPPING JSON KE EXCEL - FORM ASSET COUNTING';
    titleCell.font = { name: FONT_FAMILY, size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = headerFillNavy;
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    wsMapping.getRow(1).height = 36;

    wsMapping.getCell('A2').value = 'Dokumen acuan untuk System Analyst (SA) & Backend Developer dalam integrasi data API Asset Counting.';
    wsMapping.getCell('A2').font = { name: FONT_FAMILY, size: 10, italic: true, color: { argb: 'FF64748B' } };
    wsMapping.getRow(2).height = 20;

    // Tabel 1: Metadata Header
    wsMapping.getCell('A4').value = 'A. METADATA HEADER (Bagian Atas Laporan)';
    wsMapping.getCell('A4').font = { name: FONT_FAMILY, size: 11, bold: true, color: { argb: 'FF1E293B' } };

    const metaHeaders = ['Posisi Sel Excel', 'Teks di Excel', 'Field JSON Utama (BE)', 'Field JSON Alternatif', 'Tipe Data', 'Contoh Nilai', 'Keterangan SA/BE'];
    const rMetaHeader = wsMapping.getRow(5);
    rMetaHeader.height = 26;
    metaHeaders.forEach((h, i) => {
        const c = rMetaHeader.getCell(i + 1);
        c.value = h;
        c.fill = headerFillBlue;
        c.font = { name: FONT_FAMILY, size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = borderThin;
    });

    const metaData = [
        ['Sel A1', 'PT HASJRAT ABADI CABANG [CABANG]', 'office_branch_name', 'office_name / branch_name / cabang', 'String', 'Biak', 'Nama cabang operasional. Huruf otomatis dikonversi ke UPPERCASE.'],
        ['Sel A2', 'FORM ASSET COUNTING', '(Fixed Text)', '-', 'String', 'FORM ASSET COUNTING', 'Judul tetap laporan di baris 2.'],
        ['Sel A3', 'TAHUN BUKU [TAHUN]', 'fiscal_year', 'doc_date / periode_start / created_date', 'Integer / Date', '2025', 'Tahun buku laporan. Jika tidak dikirim, diekstrak dari tahun doc_date atau tahun berjalan.']
    ];

    metaData.forEach((row, idx) => {
        const r = wsMapping.getRow(6 + idx);
        r.height = 22;
        row.forEach((val, cIdx) => {
            const c = r.getCell(cIdx + 1);
            c.value = val;
            c.font = { name: FONT_FAMILY, size: 9 };
            c.border = borderThin;
            c.alignment = { vertical: 'middle', horizontal: cIdx === 0 ? 'center' : 'left' };
            if (idx % 2 === 1) c.fill = fillZebra;
        });
    });

    // Tabel 2: Detail Kolom Data (Array: stockDetails)
    const startRowCols = 11;
    wsMapping.getCell(`A${startRowCols - 1}`).value = 'B. DETAIL 12 KOLOM TABEL ASET (Array: stockDetails / details)';
    wsMapping.getCell(`A${startRowCols - 1}`).font = { name: FONT_FAMILY, size: 11, bold: true, color: { argb: 'FF1E293B' } };

    const colTableHeaders = ['Kolom Excel', 'Nama Kolom di Excel', 'Field JSON Utama (BE)', 'Field JSON Alternatif', 'Tipe Data', 'Format / Transformasi di Excel', 'Contoh Nilai JSON'];
    const rColHeader = wsMapping.getRow(startRowCols);
    rColHeader.height = 26;
    colTableHeaders.forEach((h, i) => {
        const c = rColHeader.getCell(i + 1);
        c.value = h;
        c.fill = headerFillBlue;
        c.font = { name: FONT_FAMILY, size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = borderThin;
    });

    const colDataMapping = [
        ['Kolom A', 'NO', 'no', 'stockdetail_id / nomor urut', 'Integer', 'Nomor urut numerik (1, 2, 3, ...)', '1'],
        ['Kolom B', 'NOMOR ASSET MODUL', 'item_code', 'item_code_sap / serial / nomor_asset_modul', 'String', 'Teks kode aset dari sistem/SAP', '30200B9902001'],
        ['Kolom C', 'NOMOR ASSET SCAN', 'serial', 'nomor_asset_scan / item_code / barcode', 'String', 'Teks barcode scan (Cell Highlight Kuning #FFFF00)', '30200B9902001'],
        ['Kolom D', 'NAMA ASSET', 'item_name', 'nama_asset / asset_name / name', 'String', 'Deskripsi/nama aset fisik', 'BANGUNAN'],
        ['Kolom E', 'TANGGAL PEROLEHAN', 'tanggal_perolehan', 'TanggalPerolehan / acquisition_date', 'Date / String', 'Format YYYY-MM-DD (timestamp jam dibuang)', '1993-10-21 00:00:00.000'],
        ['Kolom F', 'HARGA PEROLEHAN', 'harga_perolehan', 'HargaPerolehan / acquisition_cost', 'Number / String', 'Format Angka Mata Uang (#,##0)', '2696351206.000000'],
        ['Kolom G', 'AKUMULASI PENYUSUTAN', 'akumulasi_penyusutan', 'AkumulasiPenyusutan / depreciation', 'Number / String', 'Format Angka Mata Uang (#,##0). String ".000000" diubah ke 0', '.000000'],
        ['Kolom H', 'NBV', 'nbv', 'Nbv / (HargaPerolehan - AkumulasiPenyusutan)', 'Number / String', 'Net Book Value / Nilai Buku (#,##0)', '2696351206.000000'],
        ['Kolom I', 'USER/PENGGUNA', 'asset_pic', 'user_pengguna / pic / asset_location', 'String', 'Nama penanggung jawab atau lokasi fisik aset', 'Sunardi / (CAB. BIAK)'],
        ['Kolom J', 'STATUS BARANG', 'asset_condition', 'status_barang / is_checked', 'String', 'Kondisi barang (Cell Highlight Kuning #FFFF00, default: ADA)', 'ADA'],
        ['Kolom K', 'FOTO UNIT / KETERANGAN', 'attachment', 'foto_unit / img_path / foto', 'Array / String', 'Foto di-embed otomatis (1-5 thumbnail rapi)', '["https://hrms.hasjrat.co.id/horor/...jpg"]'],
        ['Kolom L', 'KETERANGAN', 'remarks', 'checked_remarks / asset_location / keterangan', 'String', 'Catatan tambahan / lokasi fisik', 'CAB. BIAK']
    ];

    colDataMapping.forEach((row, idx) => {
        const r = wsMapping.getRow(startRowCols + 1 + idx);
        r.height = 24;
        row.forEach((val, cIdx) => {
            const c = r.getCell(cIdx + 1);
            c.value = val;
            c.font = { name: FONT_FAMILY, size: 9 };
            c.border = borderThin;
            c.alignment = { vertical: 'middle', horizontal: cIdx === 0 ? 'center' : 'left', wrapText: true };
            if (idx % 2 === 1) c.fill = fillZebra;
        });
    });

    wsMapping.columns = [
        { width: 14 },
        { width: 26 },
        { width: 24 },
        { width: 34 },
        { width: 16 },
        { width: 38 },
        { width: 28 }
    ];

    // ==========================================
    // SHEET 2: CONTOH FORM ASSET COUNTING (MOCKUP)
    // ==========================================
    const wsSample = workbook.addWorksheet('2. Contoh Format Form Excel', { views: [{ showGridLines: true }] });

    // Header 1-3
    const s1 = wsSample.getCell('A1'); s1.value = 'PT HASJRAT ABADI CABANG BIAK'; s1.font = { name: FONT_FAMILY, size: 10, bold: true };
    const s2 = wsSample.getCell('A2'); s2.value = 'FORM ASSET COUNTING'; s2.font = { name: FONT_FAMILY, size: 10, bold: true };
    const s3 = wsSample.getCell('A3'); s3.value = 'TAHUN BUKU 2025'; s3.font = { name: FONT_FAMILY, size: 10, bold: true };
    wsSample.getRow(1).height = 18;
    wsSample.getRow(2).height = 18;
    wsSample.getRow(3).height = 18;
    wsSample.getRow(4).height = 10;

    // Header Tabel
    const tableHeaders = [
        { h: 'NO', w: 6, b: false },
        { h: 'NOMOR ASSET\nMODUL', w: 20, b: false },
        { h: 'NOMOR ASSET SCAN', w: 20, b: false },
        { h: 'NAMA ASSET', w: 28, b: false },
        { h: 'TANGGAL PEROLEHAN', w: 18, b: true },
        { h: 'HARGA PEROLEHAN', w: 20, b: true },
        { h: 'AKUMULASI\nPENYUSUTAN', w: 20, b: true },
        { h: 'NBV', w: 18, b: true },
        { h: 'USER/PENGGUNA', w: 22, b: false },
        { h: 'STATUS BARANG', w: 16, b: false },
        { h: 'FOTO UNIT / KETERANGAN', w: 30, b: false },
        { h: 'KETERANGAN', w: 24, b: false }
    ];

    const r5 = wsSample.getRow(5);
    r5.height = 28;
    tableHeaders.forEach((th, i) => {
        const c = r5.getCell(i + 1);
        c.value = th.h;
        c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        c.border = borderDark;
        c.font = { name: FONT_FAMILY, size: 9, bold: true };
        c.fill = th.b ? fillHeaderSoftBlue : fillHeaderYellow;
        wsSample.getColumn(i + 1).width = th.w;
    });

    const sampleRows = [
        [1, '30200B9902001', '30200B9902001', 'BANGUNAN', '1993-10-21', 2696351206, 0, 2696351206, '(CAB. BIAK)', 'ADA', '[Foto Unit Disematkan]', 'CAB. BIAK'],
        [2, '30200B1502001', '30200B1502001', 'RENOVASI BANGUNAN GEDUNG', '-', 0, 0, 0, '(CAB. BIAK)', 'ADA', '-', 'CAB. BIAK'],
        [3, '30200K0302001', '30200K0302001', 'NOUVO', '2014-07-18', 10346600, 1939987, 8406613, 'Sunardi', 'ADA', '[Foto Unit Disematkan]', 'CAB. BIAK'],
        [4, '30200K1502001', '30200K1502001', 'AVANZA (X-TARIKAN HMF)', '-', 0, 0, 0, '(CAB. BIAK)', 'ADA', '-', 'CAB. BIAK'],
        [5, '30200P0402001', '30200P0402001', '1 set Meja Negosiasi ( 4 kursi & 1 Meja )', '-', 0, 0, 0, '(CAB. BIAK)', 'ADA', '-', 'CAB. BIAK']
    ];

    sampleRows.forEach((sRow, idx) => {
        const rowNum = 6 + idx;
        const r = wsSample.getRow(rowNum);
        r.height = 24;
        sRow.forEach((val, cIdx) => {
            const c = r.getCell(cIdx + 1);
            c.value = val;
            c.border = borderDark;
            c.font = { name: FONT_FAMILY, size: 9 };
            
            // Format alignment & number
            if (cIdx === 0 || cIdx === 1 || cIdx === 2 || cIdx === 4 || cIdx === 8 || cIdx === 9 || cIdx === 10) {
                c.alignment = { horizontal: 'center', vertical: 'middle' };
            } else if (cIdx === 5 || cIdx === 6 || cIdx === 7) {
                c.alignment = { horizontal: 'right', vertical: 'middle' };
                c.numFmt = '#,##0';
            } else {
                c.alignment = { horizontal: 'left', vertical: 'middle' };
            }

            // Yellow fill for scan num and status
            if (cIdx === 2 || cIdx === 9) {
                c.fill = fillCellYellow;
            }
        });
    });

    wsSample.autoFilter = { from: { row: 5, column: 1 }, to: { row: 10, column: 12 } };

    // ==========================================
    // SHEET 3: SPESIFIKASI API & CONTOH JSON
    // ==========================================
    const wsApi = workbook.addWorksheet('3. Contoh Payload JSON API', { views: [{ showGridLines: true }] });

    wsApi.mergeCells('A1:E1');
    const apiTitle = wsApi.getCell('A1');
    apiTitle.value = 'SPESIFIKASI ENDPOINT API & CONTOH REQUEST JSON';
    apiTitle.font = { name: FONT_FAMILY, size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
    apiTitle.fill = headerFillNavy;
    apiTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    wsApi.getRow(1).height = 34;

    const apiDetails = [
        ['HTTP Method', 'POST'],
        ['Endpoint URL', 'http://<host>:8000/api/v1/generate-excel/asset-counting (atau alias: /api/v1/generate-excel/asset)'],
        ['Content-Type', 'application/json'],
        ['Response Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet (.xlsx Binary Buffer)']
    ];

    apiDetails.forEach((item, idx) => {
        const r = wsApi.getRow(3 + idx);
        r.height = 22;
        const c1 = r.getCell(1); c1.value = item[0]; c1.font = { name: FONT_FAMILY, size: 9.5, bold: true }; c1.border = borderThin;
        const c2 = r.getCell(2); c2.value = item[1]; c2.font = { name: FONT_FAMILY, size: 9.5 }; c2.border = borderThin;
        wsApi.mergeCells(`B${3 + idx}:E${3 + idx}`);
    });

    wsApi.getCell('A8').value = 'Contoh Payload JSON yang dikirimkan Backend:';
    wsApi.getCell('A8').font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF1E293B' } };

    const sampleJsonText = JSON.stringify({
        stock_id: 74,
        doc_num: "DOC-BIK-001",
        doc_date: "2025-11-18 09:03:54.000",
        office_branch_name: "Biak",
        total_item: 176,
        stockDetails: [
            {
                stockdetail_id: 2782,
                item_code: "30200B9902001",
                item_name: "BANGUNAN",
                serial: "30200B9902001",
                asset_pic: "Sunardi",
                asset_location: "CAB. BIAK",
                TanggalPerolehan: "1993-10-21 00:00:00.000",
                HargaPerolehan: "2696351206.000000",
                AkumulasiPenyusutan: ".000000",
                Nbv: "2696351206.000000",
                asset_condition: "ADA",
                attachment: [
                    "https://hrms.hasjrat.co.id/horor/hororupload/section/sample1.jpg"
                ],
                remarks: "Kondisi fisik gedung baik"
            }
        ]
    }, null, 2);

    wsApi.mergeCells('A9:E28');
    const jsonCell = wsApi.getCell('A9');
    jsonCell.value = sampleJsonText;
    jsonCell.font = { name: 'Consolas', size: 9.5, color: { argb: 'FF0F172A' } };
    jsonCell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    jsonCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    jsonCell.border = borderThin;

    wsApi.columns = [
        { width: 20 },
        { width: 30 },
        { width: 30 },
        { width: 30 },
        { width: 30 }
    ];

    const outPath = path.join(__dirname, '../Dokumentasi_Spesifikasi_Mapping_Asset_Counting.xlsx');
    await workbook.xlsx.writeFile(outPath);
    console.log('Successfully generated:', outPath);
}

createMappingDoc().catch(console.error);
