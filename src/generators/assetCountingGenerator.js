const ExcelJS = require('exceljs');
const { processCheckpointPhotos } = require('../imageHandler');

/**
 * Helper untuk parsing angka bersih dari string atau number.
 * Misal: ".000000" -> 0, "2696351206.000000" -> 2696351206, null -> 0
 */
function parseCleanNumber(val) {
    if (val === null || val === undefined || val === '') return 0;
    const str = String(val).trim();
    if (str === '' || str === '.000000' || str === '.') return 0;
    const num = Number(str);
    return isNaN(num) ? 0 : num;
}

/**
 * Normalizer untuk mengubah data JSON Asset Counting menjadi flat array item yang seragam.
 *
 * ════════════════════════════════════════════════════════════════════════
 * MAPPING FIELD (17 Kolom A–Q) — Field JSON Backend (PASTI / FIX)
 * ════════════════════════════════════════════════════════════════════════
 * A  NO                    → (auto index)
 * B  ASSET CLASS           → asset_class          (baru, request ke BE)
 * C  NOMOR SAP             → sap_number           (baru, request ke BE)
 * D  NOMOR ASSET MODUL     → item_code
 * E  NOMOR ASSET SCAN      → serial
 * F  NAMA ASSET            → item_name
 * G  TANGGAL PEROLEHAN     → tanggal_perolehan    (YYYY-MM-DD, jam dibuang)
 * H  HARGA PEROLEHAN       → harga_perolehan      (#,##0 ; null/".000000" → 0)
 * I  AKUMULASI PENYUSUTAN  → akumulasi_penyusutan (#,##0 ; null/".000000" → 0)
 * J  NBV                   → nbv                  (#,##0 ; null/".000000" → 0)
 * K  Quantity On Hand      → qty_on_hand          (baru, request ke BE ; null → 1)
 * L  Quantity Hasil Opname → qty_opname           (baru, request ke BE ; null → -)
 * M  USER/PENGGUNA         → asset_pic
 * N  STATUS BARANG         → asset_condition
 * O  KONDISI BARANG        → kondisi_barang       (baru, request ke BE)
 * P  FOTO UNIT/KETERANGAN  → attachment           (Array URL ; kosong/[] → -)
 * Q  KETERANGAN            → asset_location
 * ════════════════════════════════════════════════════════════════════════
 */
function normalizeAssetCountingData(rawPayload) {
    let itemsArray = [];
    let meta = {
        branch: 'AMBON',
        fiscal_year: new Date().getFullYear()
    };

    if (!rawPayload) return { items: [], meta };

    let inputData = rawPayload;
    if (!Array.isArray(inputData) && typeof inputData === 'object') {
        // Nama cabang
        if (inputData.office_branch_name) {
            meta.branch = inputData.office_branch_name;
        }

        // Tahun buku
        if (inputData.doc_date) {
            meta.fiscal_year = new Date(inputData.doc_date).getFullYear();
        } else if (inputData.fiscal_year) {
            meta.fiscal_year = inputData.fiscal_year;
        }

        // Array data
        if (Array.isArray(inputData.stockDetails)) {
            inputData = inputData.stockDetails;
        } else if (Array.isArray(inputData.details)) {
            inputData = inputData.details;
        } else if (Array.isArray(inputData.data)) {
            inputData = inputData.data;
        } else {
            inputData = [inputData];
        }
    }

    if (!Array.isArray(inputData)) return { items: [], meta };

    inputData.forEach((entry, idx) => {

        // ── A: NO (auto index) ──────────────────────────────────────────────
        const noVal = idx + 1;

        // ── B: ASSET CLASS → asset_class ────────────────────────────────────
        const assetClass = (entry.asset_class && String(entry.asset_class).trim())
            ? String(entry.asset_class).trim() : '-';

        // ── C: NOMOR SAP → sap_number ───────────────────────────────────────
        const nomorSap = (entry.sap_number && String(entry.sap_number).trim())
            ? String(entry.sap_number).trim() : '-';

        // ── D: NOMOR ASSET MODUL → item_code ────────────────────────────────
        const modulNum = (entry.item_code && String(entry.item_code).trim())
            ? String(entry.item_code).trim() : '-';

        // ── E: NOMOR ASSET SCAN → serial ────────────────────────────────────
        const scanNum = (entry.serial && String(entry.serial).trim())
            ? String(entry.serial).trim() : '-';

        // ── F: NAMA ASSET → item_name ────────────────────────────────────────
        const assetName = (entry.item_name && String(entry.item_name).trim())
            ? String(entry.item_name).trim() : '-';

        // ── G: TANGGAL PEROLEHAN → tanggal_perolehan ─────────────────────────
        let acqDate = '-';
        if (entry.tanggal_perolehan) {
            const cleaned = String(entry.tanggal_perolehan).split(' ')[0].split('T')[0].trim();
            acqDate = cleaned || '-';
        }

        // ── H: HARGA PEROLEHAN → harga_perolehan ─────────────────────────────
        const acqCost = parseCleanNumber(entry.harga_perolehan);

        // ── I: AKUMULASI PENYUSUTAN → akumulasi_penyusutan ───────────────────
        const accDepr = parseCleanNumber(entry.akumulasi_penyusutan);

        // ── J: NBV → nbv ─────────────────────────────────────────────────────
        const nbvVal = parseCleanNumber(entry.nbv);

        // ── K: QUANTITY ON HAND → qty_on_hand (null → 1) ─────────────────────
        const qtyOnHand = (entry.qty_on_hand !== null && entry.qty_on_hand !== undefined && entry.qty_on_hand !== '')
            ? parseCleanNumber(entry.qty_on_hand) : 1;

        // ── L: QUANTITY HASIL OPNAME → qty_opname (null → '-') ───────────────
        const qtyOpname = (entry.qty_opname !== null && entry.qty_opname !== undefined && entry.qty_opname !== '')
            ? parseCleanNumber(entry.qty_opname) : '-';

        // ── M: USER/PENGGUNA → asset_pic ─────────────────────────────────────
        const userPengguna = (entry.asset_pic && String(entry.asset_pic).trim())
            ? String(entry.asset_pic).trim() : '-';

        // ── N: STATUS BARANG → asset_condition ───────────────────────────────
        const statusBarang = (entry.asset_condition && String(entry.asset_condition).trim())
            ? String(entry.asset_condition).trim() : '-';

        // ── O: KONDISI BARANG → kondisi_barang ───────────────────────────────
        const kondisiBarang = (entry.kondisi_barang && String(entry.kondisi_barang).trim())
            ? String(entry.kondisi_barang).trim() : '-';

        // ── P: FOTO UNIT/KETERANGAN → attachment ─────────────────────────────
        const imgRef = entry.attachment || '';

        // ── Q: KETERANGAN → asset_location ───────────────────────────────────
        const keterangan = (entry.asset_location && String(entry.asset_location).trim())
            ? String(entry.asset_location).trim() : '-';

        itemsArray.push({
            no:                   noVal,
            asset_class:          assetClass,
            nomor_sap:            nomorSap,
            nomor_asset_modul:    modulNum,
            nomor_asset_scan:     scanNum,
            nama_asset:           assetName,
            tanggal_perolehan:    acqDate,
            harga_perolehan:      acqCost,
            akumulasi_penyusutan: accDepr,
            nbv:                  nbvVal,
            qty_on_hand:          qtyOnHand,
            qty_opname:           qtyOpname,
            user_pengguna:        userPengguna,
            status_barang:        statusBarang,
            kondisi_barang:       kondisiBarang,
            img_path:             imgRef,
            keterangan:           keterangan
        });
    });

    return { items: itemsArray, meta };
}

/**
 * Generator Laporan FORM ASSET COUNTING (PT HASJRAT ABADI)
 * 17 Kolom (A–Q)
 */
async function generateAssetCountingExcel(rawPayload) {
    const { items, meta } = normalizeAssetCountingData(rawPayload);

    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Form Asset Counting', {
        views: [{ showGridLines: true }]
    });

    const FONT  = 'Segoe UI';
    const TOTAL = 17;

    const fillYellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    const fillBlue   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB8CCE4' } };

    const border = {
        top:    { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left:   { style: 'thin', color: { argb: 'FF000000' } },
        right:  { style: 'thin', color: { argb: 'FF000000' } }
    };

    const aC = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const aL = { horizontal: 'left',   vertical: 'middle', wrapText: true };
    const aR = { horizontal: 'right',  vertical: 'middle', wrapText: true };

    // ── METADATA HEADER (Baris 1–3) ──────────────────────────────────────────
    const branchName = String(meta.branch || 'AMBON').toUpperCase();
    const fiscalYear  = meta.fiscal_year || new Date().getFullYear();

    const metaTexts = [
        `PT HASJRAT ABADI CABANG ${branchName}`,
        'FORM ASSET COUNTING',
        `TAHUN BUKU ${fiscalYear}`
    ];
    metaTexts.forEach((text, i) => {
        const cell = worksheet.getCell(`A${i + 1}`);
        cell.value = text;
        cell.font  = { name: FONT, size: 10, bold: true };
    });
    [1, 2, 3].forEach(r => { worksheet.getRow(r).height = 18; });
    worksheet.getRow(4).height = 10;

    // ── HEADER TABEL (Baris 5) ───────────────────────────────────────────────
    const headers = [
        { label: 'NO',                      w: 6,  fill: fillYellow },  // A
        { label: 'ASSET CLASS',             w: 18, fill: fillYellow },  // B
        { label: 'NOMOR SAP',               w: 18, fill: fillYellow },  // C
        { label: 'NOMOR ASSET\nMODUL',      w: 20, fill: fillYellow },  // D
        { label: 'NOMOR ASSET SCAN',        w: 20, fill: fillYellow },  // E
        { label: 'NAMA ASSET',              w: 28, fill: fillBlue   },  // F
        { label: 'TANGGAL\nPEROLEHAN',      w: 18, fill: fillBlue   },  // G
        { label: 'HARGA PEROLEHAN',         w: 20, fill: fillBlue   },  // H
        { label: 'AKUMULASI\nPENYUSUTAN',   w: 20, fill: fillBlue   },  // I
        { label: 'NBV',                     w: 18, fill: fillBlue   },  // J
        { label: 'Quantity On Hand',        w: 16, fill: fillYellow },  // K
        { label: 'Quantity Hasil\nOpname',  w: 16, fill: fillYellow },  // L
        { label: 'USER/PENGGUNA',           w: 22, fill: fillBlue   },  // M
        { label: 'STATUS BARANG',           w: 16, fill: fillYellow },  // N
        { label: 'KONDISI BARANG',          w: 16, fill: fillYellow },  // O
        { label: 'FOTO UNIT / KETERANGAN',  w: 65, fill: fillBlue   },  // P
        { label: 'KETERANGAN',              w: 24, fill: fillYellow }   // Q
    ];

    const HDR_ROW = 5;
    worksheet.getRow(HDR_ROW).height = 36;

    headers.forEach((h, i) => {
        const cell = worksheet.getCell(HDR_ROW, i + 1);
        cell.value     = h.label;
        cell.fill      = h.fill;
        cell.border    = border;
        cell.alignment = aC;
        cell.font      = { name: FONT, size: 9, bold: true };
        worksheet.getColumn(i + 1).width = h.w;
    });

    // ── PRE-FETCH FOTO (paralel) ─────────────────────────────────────────────
    const assetImages = await Promise.all(items.map(it => processCheckpointPhotos(it.img_path)));

    let maxFotoW = 65;
    assetImages.forEach(img => {
        if (img) {
            const w = Math.round(img.widthPx / 7) + 6;
            if (w > maxFotoW) maxFotoW = w;
        }
    });
    worksheet.getColumn(16).width = maxFotoW; // Kolom P

    // ── DATA BARIS ───────────────────────────────────────────────────────────
    let curRow = HDR_ROW;

    for (let i = 0; i < items.length; i++) {
        curRow++;
        const it  = items[i];
        const img = assetImages[i];

        const rows = [
            { v: it.no,                        a: aC },               // A
            { v: it.asset_class,               a: aC },               // B
            { v: it.nomor_sap,                 a: aC },               // C
            { v: it.nomor_asset_modul,         a: aC },               // D
            { v: it.nomor_asset_scan,          a: aC },               // E
            { v: it.nama_asset,                a: aL },               // F
            { v: it.tanggal_perolehan,         a: aC },               // G
            { v: it.harga_perolehan || 0,      a: aR, fmt: '#,##0' }, // H
            { v: it.akumulasi_penyusutan || 0, a: aR, fmt: '#,##0' }, // I
            { v: it.nbv || 0,                  a: aR, fmt: '#,##0' }, // J
            { v: it.qty_on_hand,               a: aC },               // K
            { v: it.qty_opname,                a: aC },               // L
            { v: it.user_pengguna,             a: aC },               // M
            { v: it.status_barang,             a: aC },               // N
            { v: it.kondisi_barang,            a: aC },               // O
            { v: null,                         a: aC },               // P (foto embed)
            { v: it.keterangan,                a: aL }                // Q
        ];

        rows.forEach((d, ci) => {
            const cell = worksheet.getCell(curRow, ci + 1);
            if (d.v !== null) cell.value = d.v;
            cell.alignment = d.a;
            if (d.fmt) cell.numFmt = d.fmt;
            cell.border = border;
            cell.font   = { name: FONT, size: 9 };
        });

        // Embed foto di kolom P (index 16)
        if (img) {
            const imgId = workbook.addImage({ buffer: img.buffer, extension: 'png' });
            const w     = Math.round(img.widthPx / 7) + 6;
            if (w > maxFotoW) { maxFotoW = w; worksheet.getColumn(16).width = w; }
            worksheet.getRow(curRow).height = Math.max(120, Math.round((img.heightPx + 20) * 0.75));
            worksheet.addImage(imgId, {
                tl:     { col: 15.04, row: (curRow - 1) + 0.04 },
                ext:    { width: img.widthPx, height: img.heightPx },
                editAs: 'oneCell'
            });
        } else {
            worksheet.getRow(curRow).height = 24;
            worksheet.getCell(curRow, 16).value = '-';
        }
    }

    // AutoFilter A5:Q[lastRow]
    worksheet.autoFilter = {
        from: { row: HDR_ROW, column: 1 },
        to:   { row: curRow,  column: TOTAL }
    };

    return await workbook.xlsx.writeBuffer();
}

module.exports = { generateAssetCountingExcel, normalizeAssetCountingData };
