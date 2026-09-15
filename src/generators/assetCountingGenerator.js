const ExcelJS = require('exceljs');
const { processCheckpointPhotos } = require('../imageHandler');

/**
 * Normalizer untuk mengubah data JSON Asset Counting menjadi flat array item yang seragam.
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
        if (inputData.branch_name || inputData.cabang || inputData.branch) {
            meta.branch = inputData.branch_name || inputData.cabang || inputData.branch;
        }
        if (inputData.fiscal_year || inputData.tahun_buku || inputData.tahun || inputData.year) {
            meta.fiscal_year = inputData.fiscal_year || inputData.tahun_buku || inputData.tahun || inputData.year;
        }

        if (Array.isArray(inputData.data)) {
            inputData = inputData.data;
        } else if (Array.isArray(inputData.items)) {
            inputData = inputData.items;
        } else if (Array.isArray(inputData.assets)) {
            inputData = inputData.assets;
        } else {
            inputData = [inputData];
        }
    }

    if (!Array.isArray(inputData)) return { items: [], meta };

    inputData.forEach((entry, idx) => {
        if (entry.branch_name || entry.cabang || entry.branch) {
            meta.branch = entry.branch_name || entry.cabang || entry.branch;
        }
        if (entry.fiscal_year || entry.tahun_buku || entry.tahun || entry.year) {
            meta.fiscal_year = entry.fiscal_year || entry.tahun_buku || entry.tahun || entry.year;
        }

        const noVal = entry.no || entry.nomor || (idx + 1);
        const modulNum = entry.nomor_asset_modul || entry.asset_modul || entry.asset_num || entry.asset_code || entry.kode_asset || entry.no_asset || '';
        const scanNum = entry.nomor_asset_scan || entry.asset_scan || entry.scan_num || entry.barcode || entry.asset_scan_code || '';
        const assetName = entry.nama_asset || entry.asset_name || entry.name || entry.deskripsi || '';
        const acqDate = entry.tanggal_perolehan || entry.tgl_perolehan || entry.acquisition_date || entry.date || '';
        const acqCost = Number(entry.harga_perolehan || entry.acquisition_cost || entry.cost || entry.price || 0);
        const accDepr = Number(entry.akumulasi_penyusutan || entry.accumulated_depreciation || entry.depreciation || 0);
        const nbvVal = Number(entry.nbv || entry.nilai_buku || entry.net_book_value || (acqCost - accDepr));
        const userPengguna = entry.user_pengguna || entry.pengguna || entry.user || entry.pic || entry.pemegang_asset || entry.check_user_name || '';
        const statusBarang = entry.status_barang || entry.status || entry.kondisi || 'ADA';
        const imgRef = entry.foto_unit || entry.img_path || entry.foto || entry.hasilFoto || entry.image || '';
        const keterangan = entry.keterangan || entry.notes || entry.remark || entry.description || '';

        itemsArray.push({
            no: noVal,
            nomor_asset_modul: modulNum,
            nomor_asset_scan: scanNum,
            nama_asset: assetName,
            tanggal_perolehan: acqDate,
            harga_perolehan: acqCost,
            akumulasi_penyusutan: accDepr,
            nbv: nbvVal,
            user_pengguna: userPengguna,
            status_barang: statusBarang,
            img_path: imgRef,
            keterangan: keterangan
        });
    });

    return { items: itemsArray, meta };
}

/**
 * Generator Laporan FORM ASSET COUNTING (PT HASJRAT ABADI)
 * - 12 Kolom Utama: NO, NOMOR ASSET MODUL, NOMOR ASSET SCAN, NAMA ASSET, TANGGAL PEROLEHAN,
 *   HARGA PEROLEHAN, AKUMULASI PENYUSUTAN, NBV, USER/PENGGUNA, STATUS BARANG, FOTO UNIT / KETERANGAN, KETERANGAN.
 * - Header styling presisi (Yellow & Soft Blue) sesuai screenshot.
 */
async function generateAssetCountingExcel(rawPayload) {
    const { items, meta } = normalizeAssetCountingData(rawPayload);

    const workbook = new ExcelJS.Workbook();
    const sheetTitle = 'Form Asset Counting';

    const worksheet = workbook.addWorksheet(sheetTitle, {
        views: [{ showGridLines: true }]
    });

    const FONT_FAMILY = 'Segoe UI';

    // Styling Palette Sesuai Screenshot Form Asset Counting
    const fillHeaderYellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Bright Yellow
    const fillHeaderBlue = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB8CCE4' } };   // Soft Blue
    const fillCellYellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };   // Yellow Cell Fill

    const borderCell = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
    };

    const alignCenter = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const alignLeft = { horizontal: 'left', vertical: 'middle', wrapText: true };
    const alignRight = { horizontal: 'right', vertical: 'middle', wrapText: true };

    // --- 1. METADATA HEADER (Baris 1 - 3) ---
    const branchName = String(meta.branch || 'AMBON').toUpperCase();
    const fiscalYear = meta.fiscal_year || new Date().getFullYear();

    const c1 = worksheet.getCell('A1');
    c1.value = `PT HASJRAT ABADI CABANG ${branchName}`;
    c1.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF000000' } };

    const c2 = worksheet.getCell('A2');
    c2.value = 'FORM ASSET COUNTING';
    c2.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF000000' } };

    const c3 = worksheet.getCell('A3');
    c3.value = `TAHUN BUKU ${fiscalYear}`;
    c3.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF000000' } };

    worksheet.getRow(1).height = 18;
    worksheet.getRow(2).height = 18;
    worksheet.getRow(3).height = 18;
    worksheet.getRow(4).height = 10; // Spacing row

    // --- 2. HEADER TABEL (Baris 5 - 12 Kolom) ---
    const headers = [
        { header: 'NO', key: 'no', width: 6, type: 'yellow' },
        { header: 'NOMOR ASSET\nMODUL', key: 'nomor_asset_modul', width: 20, type: 'yellow' },
        { header: 'NOMOR ASSET SCAN', key: 'nomor_asset_scan', width: 20, type: 'yellow' },
        { header: 'NAMA ASSET', key: 'nama_asset', width: 28, type: 'yellow' },
        { header: 'TANGGAL PEROLEHAN', key: 'tanggal_perolehan', width: 18, type: 'blue' },
        { header: 'HARGA PEROLEHAN', key: 'harga_perolehan', width: 20, type: 'blue' },
        { header: 'AKUMULASI\nPENYUSUTAN', key: 'akumulasi_penyusutan', width: 20, type: 'blue' },
        { header: 'NBV', key: 'nbv', width: 18, type: 'blue' },
        { header: 'USER/PENGGUNA', key: 'user_pengguna', width: 22, type: 'yellow' },
        { header: 'STATUS BARANG', key: 'status_barang', width: 16, type: 'yellow' },
        { header: 'FOTO UNIT / KETERANGAN', key: 'photo', width: 65, type: 'yellow' },
        { header: 'KETERANGAN', key: 'keterangan', width: 24, type: 'yellow' }
    ];

    const headerRowIdx = 5;
    worksheet.getRow(headerRowIdx).height = 28;

    headers.forEach((h, colIdx) => {
        const colNum = colIdx + 1;
        const cell = worksheet.getCell(headerRowIdx, colNum);
        cell.value = h.header;
        cell.alignment = alignCenter;
        cell.border = borderCell;
        cell.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: 'FF000000' } };
        worksheet.getColumn(colNum).width = h.width;

        if (h.type === 'blue') {
            cell.fill = fillHeaderBlue;
        } else {
            cell.fill = fillHeaderYellow;
        }
    });

    let currentRow = headerRowIdx;
    let maxFotoColWidth = 65;

    // Pre-fetch dan process seluruh foto secara paralel
    const assetImages = await Promise.all(
        items.map(item => processCheckpointPhotos(item.img_path))
    );

    // Hitung lebar maksimal kolom foto berdasarkan jumlah foto per baris
    assetImages.forEach(imgData => {
        if (imgData) {
            const neededColW = Math.round(imgData.widthPx / 7) + 6;
            if (neededColW > maxFotoColWidth) {
                maxFotoColWidth = neededColW;
            }
        }
    });
    worksheet.getColumn(11).width = maxFotoColWidth;

    // --- 3. DATA BARIS ---
    for (let idx = 0; idx < items.length; idx++) {
        currentRow++;
        const item = items[idx];
        const imgData = assetImages[idx];

        const rowNum = idx + 1;
        const cell1 = worksheet.getCell(currentRow, 1); cell1.value = rowNum; cell1.alignment = alignCenter;
        const cell2 = worksheet.getCell(currentRow, 2); cell2.value = item.nomor_asset_modul; cell2.alignment = alignCenter;
        const cell3 = worksheet.getCell(currentRow, 3); cell3.value = item.nomor_asset_scan; cell3.alignment = alignCenter;
        const cell4 = worksheet.getCell(currentRow, 4); cell4.value = item.nama_asset; cell4.alignment = alignLeft;
        const cell5 = worksheet.getCell(currentRow, 5); cell5.value = item.tanggal_perolehan; cell5.alignment = alignCenter;
        
        // Formatted Currency Columns
        const cell6 = worksheet.getCell(currentRow, 6); 
        cell6.value = item.harga_perolehan || 0; 
        cell6.numFmt = '#,##0'; 
        cell6.alignment = alignRight;

        const cell7 = worksheet.getCell(currentRow, 7); 
        cell7.value = item.akumulasi_penyusutan || 0; 
        cell7.numFmt = '#,##0'; 
        cell7.alignment = alignRight;

        const cell8 = worksheet.getCell(currentRow, 8); 
        cell8.value = item.nbv || 0; 
        cell8.numFmt = '#,##0'; 
        cell8.alignment = alignRight;

        const cell9 = worksheet.getCell(currentRow, 9); cell9.value = item.user_pengguna || '-'; cell9.alignment = alignCenter;
        const cell10 = worksheet.getCell(currentRow, 10); cell10.value = item.status_barang || 'ADA'; cell10.alignment = alignCenter;
        const cell11 = worksheet.getCell(currentRow, 11); cell11.alignment = alignCenter;
        const cell12 = worksheet.getCell(currentRow, 12); cell12.value = item.keterangan || ''; cell12.alignment = alignLeft;

        // Apply borders & font styling
        for (let c = 1; c <= 12; c++) {
            const cell = worksheet.getCell(currentRow, c);
            cell.border = borderCell;
            cell.font = { name: FONT_FAMILY, size: 9, color: { argb: 'FF000000' } };
        }

        // Kolom C (NOMOR ASSET SCAN) & Kolom J (STATUS BARANG) memiliki latar kuning sesuai screenshot
        cell3.fill = fillCellYellow;
        cell10.fill = fillCellYellow;

        // Handle Photo Embedding di Kolom 11 (K)
        if (imgData) {
            const imageId = workbook.addImage({
                buffer: imgData.buffer,
                extension: 'png'
            });

            const neededColW = Math.round(imgData.widthPx / 7) + 6;
            if (neededColW > maxFotoColWidth) {
                maxFotoColWidth = neededColW;
                worksheet.getColumn(11).width = maxFotoColWidth;
            }

            const calcHeight = Math.max(120, Math.round((imgData.heightPx + 20) * 0.75));
            worksheet.getRow(currentRow).height = calcHeight;

            const rZero = currentRow - 1;
            worksheet.addImage(imageId, {
                tl: { col: 10.04, row: rZero + 0.04 }, // 0-indexed column 10 = Column K
                ext: { width: imgData.widthPx, height: imgData.heightPx },
                editAs: 'oneCell'
            });
        } else {
            worksheet.getRow(currentRow).height = 24;
            cell11.value = '-';
        }
    }

    // AutoFilter Excel pada baris 5 (A5:L[lastRow])
    worksheet.autoFilter = {
        from: { row: headerRowIdx, column: 1 },
        to: { row: currentRow, column: 12 }
    };

    return await workbook.xlsx.writeBuffer();
}

module.exports = {
    generateAssetCountingExcel,
    normalizeAssetCountingData
};
