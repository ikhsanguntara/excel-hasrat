const ExcelJS = require('exceljs');
const { processCheckpointPhotos } = require('../imageHandler');

/**
 * Helper untuk parsing angka bersih dari string atau number (misal ".000000" -> 0)
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
 * MAPPING FIELD (17 Kolom):
 * A  NO                   -> auto index
 * B  ASSET CLASS          -> asset_class | category | item_category
 * C  NOMOR SAP            -> sap_number | sap_code | item_sap | nomor_sap
 * D  NOMOR ASSET MODUL    -> item_code | nomor_asset_modul
 * E  NOMOR ASSET SCAN     -> serial | nomor_asset_scan | item_code (fallback)
 * F  NAMA ASSET           -> item_name | nama_asset
 * G  TANGGAL PEROLEHAN    -> tanggal_perolehan | TanggalPerolehan
 * H  HARGA PEROLEHAN      -> harga_perolehan | HargaPerolehan
 * I  AKUMULASI PENYUSUTAN -> akumulasi_penyusutan | AkumulasiPenyusutan
 * J  NBV                  -> nbv | Nbv  (fallback: H - I)
 * K  QUANTITY ON HAND     -> qty_on_hand | quantity_on_hand | qty  (null/kosong -> 1)
 * L  QUANTITY HASIL OPNAME-> qty_opname | quantity_opname | qty_hasil_opname  (null -> '-')
 * M  USER/PENGGUNA        -> asset_pic | user_pengguna
 * N  STATUS BARANG        -> asset_condition | status_barang
 * O  KONDISI BARANG       -> kondisi_barang | asset_status | condition
 * P  FOTO UNIT/KETERANGAN -> attachment | img_path
 * Q  KETERANGAN           -> remarks | keterangan | asset_location
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
        // Cabang
        if (inputData.office_branch_name) {
            meta.branch = inputData.office_branch_name;
        } else if (inputData.office_name || inputData.branch_name) {
            meta.branch = inputData.office_name || inputData.branch_name;
        }

        // Tahun Buku
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
        // A: NO
        const noVal = idx + 1;

        // B: ASSET CLASS
        const assetClass = entry.asset_class || entry.category || entry.item_category || '-';

        // C: NOMOR SAP
        const nomorSap = entry.sap_number || entry.sap_code || entry.item_sap || entry.nomor_sap || '-';

        // D: NOMOR ASSET MODUL
        const modulNum = entry.item_code || entry.nomor_asset_modul || '-';

        // E: NOMOR ASSET SCAN
        const scanNum = entry.serial || entry.nomor_asset_scan || entry.item_code || '-';

        // F: NAMA ASSET
        const assetName = entry.item_name || entry.nama_asset || '-';

        // G: TANGGAL PEROLEHAN
        let acqDate = '-';
        const rawDate = entry.tanggal_perolehan || entry.TanggalPerolehan || '';
        if (rawDate) {
            const cleaned = String(rawDate).split(' ')[0].split('T')[0].trim();
            acqDate = cleaned || '-';
        }

        // H: HARGA PEROLEHAN
        const acqCost = parseCleanNumber(entry.harga_perolehan !== undefined ? entry.harga_perolehan : entry.HargaPerolehan);

        // I: AKUMULASI PENYUSUTAN
        const accDepr = parseCleanNumber(entry.akumulasi_penyusutan !== undefined ? entry.akumulasi_penyusutan : entry.AkumulasiPenyusutan);

        // J: NBV
        let nbvVal = 0;
        const rawNbv = entry.nbv !== undefined ? entry.nbv : entry.Nbv;
        if (rawNbv !== undefined && rawNbv !== null && rawNbv !== '') {
            nbvVal = parseCleanNumber(rawNbv);
        } else {
            nbvVal = acqCost - accDepr;
        }

        // K: QUANTITY ON HAND (null/kosong -> default 1)
        const rawQtyOnHand = entry.qty_on_hand !== undefined ? entry.qty_on_hand
                           : entry.quantity_on_hand !== undefined ? entry.quantity_on_hand
                           : entry.qty;
        const qtyOnHand = (rawQtyOnHand !== null && rawQtyOnHand !== undefined && rawQtyOnHand !== '')
            ? parseCleanNumber(rawQtyOnHand) : 1;

        // L: QUANTITY HASIL OPNAME (null/kosong -> '-')
        const rawQtyOpname = entry.qty_opname !== undefined ? entry.qty_opname
                           : entry.quantity_opname !== undefined ? entry.quantity_opname
                           : entry.qty_hasil_opname;
        const qtyOpname = (rawQtyOpname !== null && rawQtyOpname !== undefined && rawQtyOpname !== '')
            ? parseCleanNumber(rawQtyOpname) : '-';

        // M: USER/PENGGUNA
        const userPengguna = (entry.asset_pic && String(entry.asset_pic).trim())
            ? String(entry.asset_pic).trim()
            : (entry.user_pengguna && String(entry.user_pengguna).trim())
                ? String(entry.user_pengguna).trim() : '-';

        // N: STATUS BARANG
        const statusBarang = (entry.asset_condition && String(entry.asset_condition).trim())
            ? String(entry.asset_condition).trim()
            : (entry.status_barang && String(entry.status_barang).trim())
                ? String(entry.status_barang).trim() : '-';

        // O: KONDISI BARANG
        const kondisiBarang = (entry.kondisi_barang && String(entry.kondisi_barang).trim())
            ? String(entry.kondisi_barang).trim()
            : (entry.asset_status && String(entry.asset_status).trim())
                ? String(entry.asset_status).trim()
                : (entry.condition && String(entry.condition).trim())
                    ? String(entry.condition).trim() : '-';

        // P: FOTO UNIT / KETERANGAN
        const imgRef = entry.attachment || entry.img_path || '';

        // Q: KETERANGAN
        const keterangan = (entry.remarks && String(entry.remarks).trim())
            ? String(entry.remarks).trim()
            : (entry.keterangan && String(entry.keterangan).trim())
                ? String(entry.keterangan).trim()
                : (entry.asset_location || '-');

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
 * 17 Kolom (A-Q): NO, ASSET CLASS, NOMOR SAP, NOMOR ASSET MODUL, NOMOR ASSET SCAN,
 * NAMA ASSET, TANGGAL PEROLEHAN, HARGA PEROLEHAN, AKUMULASI PENYUSUTAN, NBV,
 * Quantity On Hand, Quantity Hasil Opname, USER/PENGGUNA, STATUS BARANG,
 * KONDISI BARANG, FOTO UNIT/KETERANGAN, KETERANGAN
 */
async function generateAssetCountingExcel(rawPayload) {
    const { items, meta } = normalizeAssetCountingData(rawPayload);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Form Asset Counting', {
        views: [{ showGridLines: true }]
    });

    const FONT  = 'Segoe UI';
    const TOTAL = 17; // total kolom

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

    // ── METADATA HEADER (Baris 1-4) ─────────────────────────────────────────
    const branchName = String(meta.branch || 'AMBON').toUpperCase();
    const fiscalYear  = meta.fiscal_year || new Date().getFullYear();

    ['A1', 'A2', 'A3'].forEach((addr, i) => {
        const cell = worksheet.getCell(addr);
        cell.value = [`PT HASJRAT ABADI CABANG ${branchName}`, 'FORM ASSET COUNTING', `TAHUN BUKU ${fiscalYear}`][i];
        cell.font  = { name: FONT, size: 10, bold: true };
    });
    [1, 2, 3].forEach(r => { worksheet.getRow(r).height = 18; });
    worksheet.getRow(4).height = 10;

    // ── HEADER TABEL (Baris 5) ───────────────────────────────────────────────
    const headers = [
        { label: 'NO',                      w: 6,  fill: fillYellow },
        { label: 'ASSET CLASS',             w: 18, fill: fillYellow },
        { label: 'NOMOR SAP',               w: 18, fill: fillYellow },
        { label: 'NOMOR ASSET\nMODUL',      w: 20, fill: fillYellow },
        { label: 'NOMOR ASSET SCAN',        w: 20, fill: fillYellow },
        { label: 'NAMA ASSET',              w: 28, fill: fillBlue   },
        { label: 'TANGGAL\nPEROLEHAN',      w: 18, fill: fillBlue   },
        { label: 'HARGA PEROLEHAN',         w: 20, fill: fillBlue   },
        { label: 'AKUMULASI\nPENYUSUTAN',   w: 20, fill: fillBlue   },
        { label: 'NBV',                     w: 18, fill: fillBlue   },
        { label: 'Quantity On Hand',        w: 16, fill: fillYellow },
        { label: 'Quantity Hasil\nOpname',  w: 16, fill: fillYellow },
        { label: 'USER/PENGGUNA',           w: 22, fill: fillBlue   },
        { label: 'STATUS BARANG',           w: 16, fill: fillBlue   },
        { label: 'KONDISI BARANG',          w: 16, fill: fillYellow },
        { label: 'FOTO UNIT / KETERANGAN',  w: 65, fill: fillBlue   },
        { label: 'KETERANGAN',              w: 24, fill: fillYellow }
    ];

    const HDR_ROW = 5;
    worksheet.getRow(HDR_ROW).height = 36;

    headers.forEach((h, i) => {
        const col  = i + 1;
        const cell = worksheet.getCell(HDR_ROW, col);
        cell.value     = h.label;
        cell.fill      = h.fill;
        cell.border    = border;
        cell.alignment = aC;
        cell.font      = { name: FONT, size: 9, bold: true };
        worksheet.getColumn(col).width = h.w;
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

        const values = [
            { v: it.no,                   a: aC },                                   // A
            { v: it.asset_class,          a: aC },                                   // B
            { v: it.nomor_sap,            a: aC },                                   // C
            { v: it.nomor_asset_modul,    a: aC },                                   // D
            { v: it.nomor_asset_scan,     a: aC },                                   // E
            { v: it.nama_asset,           a: aL },                                   // F
            { v: it.tanggal_perolehan,    a: aC },                                   // G
            { v: it.harga_perolehan || 0, a: aR, fmt: '#,##0' },                    // H
            { v: it.akumulasi_penyusutan || 0, a: aR, fmt: '#,##0' },              // I
            { v: it.nbv || 0,             a: aR, fmt: '#,##0' },                    // J
            { v: it.qty_on_hand,          a: aC },                                   // K
            { v: it.qty_opname,           a: aC },                                   // L
            { v: it.user_pengguna,        a: aC },                                   // M
            { v: it.status_barang,        a: aC },                                   // N
            { v: it.kondisi_barang,       a: aC },                                   // O
            { v: null,                    a: aC },                                   // P (foto)
            { v: it.keterangan,           a: aL }                                    // Q
        ];

        values.forEach((d, ci) => {
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
            const w = Math.round(img.widthPx / 7) + 6;
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
