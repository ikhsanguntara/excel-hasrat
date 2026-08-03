const ExcelJS = require('exceljs');
const { processCheckpointPhotos } = require('../imageHandler');

/**
 * Generator Laporan Checkpoint YAMAHA (1 Row Flat Table tanpa Merged Banner)
 * - 8 Kolom Lengkap: No, Area, Section, Sub-bagian, Pertanyaan Checkpoint, Hasil, Waktu Cek, Foto Lampiran.
 * - Foto berukuran Besar & Wide (220x150 px) tanpa gepeng/terdistorsi.
 */
async function generateYamahaExcel(items, areaFilter = null, sectionFilter = null, subdetailFilter = null) {
    let filteredItems = items;

    // 1. Filter per Area jika ditentukan
    if (areaFilter && areaFilter !== 'ALL' && areaFilter !== 'Semua Area') {
        filteredItems = filteredItems.filter(item => {
            const area = String(item.area_name || '').trim().toLowerCase();
            return area === String(areaFilter).trim().toLowerCase();
        });
    }

    // 2. Filter per Section jika ditentukan
    if (sectionFilter && sectionFilter !== 'ALL' && sectionFilter !== 'Semua Section') {
        filteredItems = filteredItems.filter(item => {
            const sec = String(item.section_name || '').trim().toLowerCase();
            return sec === String(sectionFilter).trim().toLowerCase();
        });
    }

    // 3. Filter per Sub-bagian jika ditentukan
    if (subdetailFilter && subdetailFilter !== 'ALL' && subdetailFilter !== 'Semua Sub-bagian') {
        filteredItems = filteredItems.filter(item => {
            const sub = String(item.sectiondtl_name || '').trim().toLowerCase();
            return sub === String(subdetailFilter).trim().toLowerCase();
        });
    }

    if (!filteredItems || filteredItems.length === 0) {
        filteredItems = items; // Fallback
    }

    const workbook = new ExcelJS.Workbook();
    const sheetTitle = 'Laporan Checkpoint Yamaha';

    const worksheet = workbook.addWorksheet(sheetTitle, {
        views: [{ showGridLines: true }]
    });

    const FONT_FAMILY = 'Segoe UI';

    const fillHeaderGray = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };  // Header Light Gray
    const fillTitle = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    const fillZebra = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };

    const borderCell = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };

    const alignCenter = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const alignLeft = { horizontal: 'left', vertical: 'middle', wrapText: true };

    // --- 1. BANNER JUDUL YAMAHA (Merged A1:H1) ---
    worksheet.mergeCells('A1:H1');
    const titleCell = worksheet.getCell('A1');
    
    let filterSubtitle = [];
    if (areaFilter && areaFilter !== 'ALL') filterSubtitle.push(`Area: ${areaFilter}`);
    if (sectionFilter && sectionFilter !== 'ALL') filterSubtitle.push(`Sec: ${sectionFilter}`);
    if (subdetailFilter && subdetailFilter !== 'ALL') filterSubtitle.push(`Sub: ${subdetailFilter}`);
    
    const bannerText = filterSubtitle.length > 0
        ? `  List Kondisi Fasilitas Cabang - YAMAHA [${filterSubtitle.join(' | ')}]`
        : '  List Kondisi Fasilitas Cabang - YAMAHA';

    titleCell.value = bannerText;
    titleCell.font = { name: FONT_FAMILY, size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = fillTitle;
    titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.getRow(1).height = 36;

    // --- 2. HEADER RINGKASAN METADATA DOKUMEN ---
    const firstItem = (filteredItems && filteredItems.length > 0) ? filteredItems[0] : {};
    const docNum = firstItem.doc_num || '-';
    const docDate = firstItem.doc_date || '-';
    const checkUser = firstItem.check_user || '-';
    const productGroup = firstItem.product_group || 'YAMAHA';
    const startDate = firstItem.start_doc_date || '-';
    const endDate = firstItem.end_doc_date || '-';

    const metaInfo = [
        ['No. Dokumen', docNum, 'Product Group', productGroup],
        ['Tanggal Dokumen', docDate, 'Petugas Pemeriksa', checkUser],
        ['Periode Inspeksi', `${startDate} s/d ${endDate}`, 'Total Checkpoint', `${filteredItems.length} Item`]
    ];

    metaInfo.forEach((row, idx) => {
        const rowIdx = idx + 3;
        worksheet.getRow(rowIdx).height = 18;

        const cL1 = worksheet.getCell(`A${rowIdx}`);
        cL1.value = row[0];
        cL1.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: 'FF475569' } };
        cL1.alignment = alignLeft;

        const cV1 = worksheet.getCell(`B${rowIdx}`);
        cV1.value = row[1];
        cV1.font = { name: FONT_FAMILY, size: 9, color: { argb: 'FF0F172A' } };
        cV1.alignment = alignLeft;

        const cL2 = worksheet.getCell(`D${rowIdx}`);
        cL2.value = row[2];
        cL2.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: 'FF475569' } };
        cL2.alignment = alignLeft;

        const cV2 = worksheet.getCell(`E${rowIdx}`);
        cV2.value = row[3];
        cV2.font = { name: FONT_FAMILY, size: 9, color: { argb: 'FF0F172A' } };
        cV2.alignment = alignLeft;
    });

    let currentRow = 7;

    // --- 3. HEADER TABEL DATA YAMAHA (8 KOLOM MURNI) ---
    const headers = [
        { header: 'No', key: 'no', width: 6 },
        { header: 'Area', key: 'area', width: 22 },
        { header: 'Section', key: 'section', width: 20 },
        { header: 'Sub-bagian', key: 'subdetail', width: 24 },
        { header: 'Pertanyaan Checkpoint', key: 'question', width: 44 },
        { header: 'Hasil', key: 'result', width: 14 },
        { header: 'Waktu Cek', key: 'date', width: 20 },
        { header: 'Foto Lampiran', key: 'photo', width: 75 }
    ];

    worksheet.getRow(currentRow).height = 26;
    headers.forEach((h, colIdx) => {
        const colNum = colIdx + 1;
        const cell = worksheet.getCell(currentRow, colNum);
        cell.value = h.header;
        cell.fill = fillHeaderGray;
        cell.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF1E293B' } };
        cell.alignment = alignCenter;
        cell.border = borderCell;
        worksheet.getColumn(colNum).width = h.width;
    });

    const headerRowIdx = currentRow;

    let rowCounter = 1;
    let maxFotoColWidth = 75;

    for (let idx = 0; idx < filteredItems.length; idx++) {
        currentRow++;
        const item = filteredItems[idx];

        const areaName = item.area_name || 'Area Tidak Terdefinisi';
        const sectionName = item.section_name || 'Section Umum';
        const subDtlName = item.sectiondtl_name || '-';
        const cpName = item.checkpoint_name || '';
        const resultVal = String(item.result || '').trim();
        const secDate = item.section_date || '';
        const imgPathStr = item.img_path || '';

        const isEven = (rowCounter % 2 === 0);

        // Set Cell Values (8 Kolom)
        const c1 = worksheet.getCell(currentRow, 1); c1.value = rowCounter; c1.alignment = alignCenter;
        const c2 = worksheet.getCell(currentRow, 2); c2.value = areaName; c2.alignment = alignCenter;
        const c3 = worksheet.getCell(currentRow, 3); c3.value = sectionName; c3.alignment = alignCenter;
        const c4 = worksheet.getCell(currentRow, 4); c4.value = subDtlName; c4.alignment = alignCenter;
        const c5 = worksheet.getCell(currentRow, 5); c5.value = cpName; c5.alignment = alignLeft;
        const c6 = worksheet.getCell(currentRow, 6); c6.value = resultVal; c6.alignment = alignCenter;
        const c7 = worksheet.getCell(currentRow, 7); c7.value = secDate; c7.alignment = alignCenter;
        const c8 = worksheet.getCell(currentRow, 8); c8.alignment = alignCenter;

        // Color Result Badge
        if (resultVal.toUpperCase() === 'Y') {
            c6.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
            c6.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF15803D' } };
        } else if (resultVal.toUpperCase() === 'N') {
            c6.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
            c6.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FFB91C1C' } };
        } else {
            c6.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
            c6.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF0369A1' } };
        }

        // Apply borders & font default
        for (let c = 1; c <= 8; c++) {
            const cell = worksheet.getCell(currentRow, c);
            cell.border = borderCell;
            if (c !== 6) {
                cell.font = { name: FONT_FAMILY, size: 9.5, color: { argb: 'FF334155' } };
                if (isEven) cell.fill = fillZebra;
            }
        }

        // Process Photos ke Kolom 8 (H -> 0-index 7)
        const imgData = await processCheckpointPhotos(imgPathStr);
        if (imgData) {
            const imageId = workbook.addImage({
                buffer: imgData.buffer,
                extension: 'png'
            });

            const neededColW = Math.round(imgData.widthPx / 7) + 6;
            if (neededColW > maxFotoColWidth) {
                maxFotoColWidth = neededColW;
            }
            worksheet.getColumn(8).width = maxFotoColWidth;

            const calcHeight = Math.max(130, Math.round((imgData.heightPx + 20) * 0.75));
            worksheet.getRow(currentRow).height = calcHeight;

            const rZero = currentRow - 1;
            worksheet.addImage(imageId, {
                tl: { col: 7.04, row: rZero + 0.04 },
                ext: { width: imgData.widthPx, height: imgData.heightPx },
                editAs: 'oneCell'
            });
        } else {
            worksheet.getRow(currentRow).height = 28;
            c8.value = '-';
        }

        rowCounter++;
    }

    // AUTOFILTER EXCEL AKTIF PADA RANGE A7:H[lastRow]
    worksheet.autoFilter = {
        from: { row: headerRowIdx, column: 1 },
        to: { row: currentRow, column: 8 }
    };

    return await workbook.xlsx.writeBuffer();
}

module.exports = {
    generateYamahaExcel
};
