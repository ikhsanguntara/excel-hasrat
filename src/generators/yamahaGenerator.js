const ExcelJS = require('exceljs');
const { processCheckpointPhotos } = require('../imageHandler');

/**
 * Helper untuk membersihkan nama worksheet agar sesuai aturan Excel:
 * Maksimal 31 karakter & tidak boleh mengandung karakter khusus \ / ? * : [ ]
 */
function sanitizeSheetName(name) {
    if (!name) return 'Area';
    let clean = String(name).replace(/[/\\?%*:|[\]]/g, '_').trim();
    if (clean.length > 30) {
        clean = clean.substring(0, 30);
    }
    return clean;
}

/**
 * Membangun 1 Worksheet Excel lengkap dengan header, metadata, grouping, dan foto.
 */
async function buildAreaSheet(workbook, sheetTitle, items) {
    const worksheet = workbook.addWorksheet(sheetTitle, {
        views: [{ showGridLines: true }]
    });

    const FONT_FAMILY = 'Segoe UI';

    // Palet Warna Slate Executive
    const fillTitle = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    const fillTh = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
    const fillArea = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF475569' } };
    const fillSection = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF64748B' } };
    const fillSubdetail = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    const fillZebra = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };

    const borderCell = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };

    const alignCenter = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const alignLeft = { horizontal: 'left', vertical: 'middle', wrapText: true };

    // --- 1. BANNER JUDUL YAMAHA (A1:E1) ---
    worksheet.mergeCells('A1:E1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `  LAPORAN HASIL CHECKPOINT & INSPEKSI (YAMAHA) - ${sheetTitle.toUpperCase()}`;
    titleCell.font = { name: FONT_FAMILY, size: 15, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = fillTitle;
    titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.getRow(1).height = 40;

    // --- 2. HEADER RINGKASAN METADATA DOKUMEN ---
    const firstItem = (items && items.length > 0) ? items[0] : {};
    const docNum = firstItem.doc_num || '-';
    const docDate = firstItem.doc_date || '-';
    const checkUser = firstItem.check_user || '-';
    const productGroup = firstItem.product_group || 'YAMAHA';
    const startDate = firstItem.start_doc_date || '-';
    const endDate = firstItem.end_doc_date || '-';

    const metaInfo = [
        ['No. Dokumen', docNum, 'Product Group', productGroup],
        ['Tanggal Dokumen', docDate, 'Petugas Pemeriksa', checkUser],
        ['Periode Inspeksi', `${startDate} s/d ${endDate}`, 'Total Checkpoint', `${items.length} Item`]
    ];

    metaInfo.forEach((row, idx) => {
        const rowIdx = idx + 3;
        worksheet.getRow(rowIdx).height = 20;

        const cL1 = worksheet.getCell(`A${rowIdx}`);
        cL1.value = row[0];
        cL1.font = { name: FONT_FAMILY, size: 9.5, bold: true, color: { argb: 'FF475569' } };
        cL1.alignment = alignLeft;

        const cV1 = worksheet.getCell(`B${rowIdx}`);
        cV1.value = row[1];
        cV1.font = { name: FONT_FAMILY, size: 9.5, color: { argb: 'FF0F172A' } };
        cV1.alignment = alignLeft;

        const cL2 = worksheet.getCell(`C${rowIdx}`);
        cL2.value = row[2];
        cL2.font = { name: FONT_FAMILY, size: 9.5, bold: true, color: { argb: 'FF475569' } };
        cL2.alignment = alignLeft;

        const cV2 = worksheet.getCell(`D${rowIdx}`);
        cV2.value = row[3];
        cV2.font = { name: FONT_FAMILY, size: 9.5, color: { argb: 'FF0F172A' } };
        cV2.alignment = alignLeft;
    });

    let currentRow = 7;

    // --- 3. HEADER TABEL DATA (5 KOLOM) ---
    const headers = [
        { header: 'No', key: 'no', width: 6 },
        { header: 'Pertanyaan Checkpoint', key: 'question', width: 52 },
        { header: 'Hasil', key: 'result', width: 14 },
        { header: 'Waktu Cek', key: 'date', width: 22 },
        { header: 'Foto Lampiran', key: 'photo', width: 55 }
    ];

    worksheet.getRow(currentRow).height = 28;
    headers.forEach((h, colIdx) => {
        const colNum = colIdx + 1;
        const cell = worksheet.getCell(currentRow, colNum);
        cell.value = h.header;
        cell.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = fillTh;
        cell.alignment = alignCenter;
        cell.border = borderCell;
        worksheet.getColumn(colNum).width = h.width;
    });

    // Aktifkan AutoFilter pada header tabel
    worksheet.autoFilter = {
        from: { row: currentRow, column: 1 },
        to: { row: currentRow, column: 5 }
    };

    currentRow++;

    // --- 4. GROUPING DATA PER AREA & SECTION ---
    const groupedData = {};
    items.forEach(item => {
        const area = item.area_name || 'Area Tidak Terdefinisi';
        const sec = item.section_name || 'Section Umum';
        if (!groupedData[area]) groupedData[area] = {};
        if (!groupedData[area][sec]) groupedData[area][sec] = [];
        groupedData[area][sec].push(item);
    });

    let rowCounter = 1;
    let maxFotoColWidth = 55;

    for (const [areaName, sections] of Object.entries(groupedData)) {
        // HEADER AREA (A:E)
        worksheet.getRow(currentRow).height = 24;
        worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
        const areaCell = worksheet.getCell(`A${currentRow}`);
        areaCell.value = ` AREA: ${areaName.toUpperCase()}`;
        areaCell.font = { name: FONT_FAMILY, size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        areaCell.fill = fillArea;
        areaCell.alignment = { horizontal: 'left', vertical: 'middle' };
        for (let c = 1; c <= 5; c++) worksheet.getCell(currentRow, c).border = borderCell;
        currentRow++;

        for (const [sectionName, secItems] of Object.entries(sections)) {
            // HEADER SECTION (A:E)
            worksheet.getRow(currentRow).height = 22;
            worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
            const secCell = worksheet.getCell(`A${currentRow}`);
            secCell.value = `   Section: ${sectionName}`;
            secCell.font = { name: FONT_FAMILY, size: 10.5, bold: true, color: { argb: 'FFFFFFFF' } };
            secCell.fill = fillSection;
            secCell.alignment = { horizontal: 'left', vertical: 'middle' };
            for (let c = 1; c <= 5; c++) worksheet.getCell(currentRow, c).border = borderCell;
            currentRow++;

            let lastSecDtl = null;
            for (const item of secItems) {
                const secDtl = item.sectiondtl_name;
                if (secDtl && secDtl !== lastSecDtl) {
                    worksheet.getRow(currentRow).height = 20;
                    worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
                    const dtlCell = worksheet.getCell(`A${currentRow}`);
                    dtlCell.value = `     • Sub-bagian: ${secDtl}`;
                    dtlCell.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF1E293B' } };
                    dtlCell.fill = fillSubdetail;
                    dtlCell.alignment = { horizontal: 'left', vertical: 'middle' };
                    for (let c = 1; c <= 5; c++) worksheet.getCell(currentRow, c).border = borderCell;
                    currentRow++;
                    lastSecDtl = secDtl;
                }

                const cpName = item.checkpoint_name || '';
                const resultVal = String(item.result || '').trim();
                const secDate = item.section_date || '';
                const imgPathStr = item.img_path || '';

                const isEven = (rowCounter % 2 === 0);

                // Set Cell Values (5 Kolom)
                const c1 = worksheet.getCell(currentRow, 1); c1.value = rowCounter; c1.alignment = alignCenter;
                const c2 = worksheet.getCell(currentRow, 2); c2.value = cpName; c2.alignment = alignLeft;
                const c3 = worksheet.getCell(currentRow, 3); c3.value = resultVal; c3.alignment = alignCenter;
                const c4 = worksheet.getCell(currentRow, 4); c4.value = secDate; c4.alignment = alignCenter;
                const c5 = worksheet.getCell(currentRow, 5); c5.alignment = alignCenter;

                // Color Result Badge
                if (resultVal.toUpperCase() === 'Y') {
                    c3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
                    c3.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF15803D' } };
                } else if (resultVal.toUpperCase() === 'N') {
                    c3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
                    c3.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FFB91C1C' } };
                } else {
                    c3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
                    c3.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF0369A1' } };
                }

                // Apply borders & font default
                for (let c = 1; c <= 5; c++) {
                    const cell = worksheet.getCell(currentRow, c);
                    cell.border = borderCell;
                    if (c !== 3) {
                        cell.font = { name: FONT_FAMILY, size: 9.5, color: { argb: 'FF334155' } };
                        if (isEven) cell.fill = fillZebra;
                    }
                }

                // Process Photos
                const imgData = await processCheckpointPhotos(imgPathStr);
                if (imgData) {
                    const imageId = workbook.addImage({
                        buffer: imgData.buffer,
                        extension: 'png'
                    });

                    worksheet.addImage(imageId, {
                        tl: { col: 4.08, row: currentRow - 0.9 },
                        ext: { width: imgData.widthPx, height: imgData.heightPx },
                        editAs: 'oneCell'
                    });

                    const calcHeight = Math.max(85, Math.round((imgData.heightPx + 16) * 0.75));
                    worksheet.getRow(currentRow).height = calcHeight;

                    const neededColW = Math.round(imgData.widthPx / 7) + 4;
                    if (neededColW > maxFotoColWidth) {
                        maxFotoColWidth = neededColW;
                        worksheet.getColumn(5).width = maxFotoColWidth;
                    }
                } else {
                    worksheet.getRow(currentRow).height = 28;
                    c5.value = '-';
                }

                rowCounter++;
                currentRow++;
            }
        }
    }
}

/**
 * Main Generator Laporan Checkpoint YAMAHA
 * Memikulkan pembuatan Master Sheet ("Semua Area") DAN Sheet Tab terpisah per masing-masing Area!
 */
async function generateYamahaExcel(items, areaFilter = null) {
    const workbook = new ExcelJS.Workbook();

    // Grouping item per area
    const areaGroups = {};
    items.forEach(item => {
        const areaName = item.area_name || 'Area Lainnya';
        if (!areaGroups[areaName]) areaGroups[areaName] = [];
        areaGroups[areaName].push(item);
    });

    const uniqueAreaNames = Object.keys(areaGroups);

    // 1. Buat Sheet Utama: "Semua Area" (jika areaFilter tidak dibatasi ke 1 area)
    if (!areaFilter || areaFilter === 'ALL' || areaFilter === 'Semua Area') {
        await buildAreaSheet(workbook, 'Semua Area', items);
    }

    // 2. Buat Sheet Tab terpisah per masing-masing Area!
    // Ini memudahkan pengguna mengeklik Tab Area di bagian bawah Excel
    for (const areaName of uniqueAreaNames) {
        // Jika ada areaFilter khusus (misal user pilih "Area Showroom"), buat sheet area tersebut
        if (areaFilter && areaFilter !== 'ALL' && areaFilter !== 'Semua Area') {
            if (String(areaName).trim().toLowerCase() !== String(areaFilter).trim().toLowerCase()) {
                continue; // Skip area yang tidak dipilih
            }
        }

        const sheetTitle = sanitizeSheetName(areaName);
        const areaItems = areaGroups[areaName];
        await buildAreaSheet(workbook, sheetTitle, areaItems);
    }

    // Fallback jika tidak ada sheet yang terbentuk
    if (workbook.worksheets.length === 0) {
        await buildAreaSheet(workbook, 'Semua Area', items);
    }

    return await workbook.xlsx.writeBuffer();
}

module.exports = {
    generateYamahaExcel
};
