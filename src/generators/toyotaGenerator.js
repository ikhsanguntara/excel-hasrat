const ExcelJS = require('exceljs');
const { processCheckpointPhotos } = require('../imageHandler');

/**
 * Normalizer untuk mengubah data JSON Toyota (nested area -> sections -> items)
 * menjadi flat array item agar mudah di-render per baris.
 */
function normalizeToyotaData(rawPayload) {
    let itemsArray = [];

    if (!rawPayload) return [];

    let inputData = rawPayload;
    if (!Array.isArray(inputData) && typeof inputData === 'object') {
        if (Array.isArray(inputData.data)) {
            inputData = inputData.data;
        } else {
            inputData = [inputData];
        }
    }

    if (!Array.isArray(inputData)) return [];

    inputData.forEach(entry => {
        // Jika sudah berupa flat item
        if (entry.checkpoint_name || (entry.checkPoint && !entry.sections)) {
            itemsArray.push({
                area_name: entry.area_name || entry.name || 'Showroom Toyota',
                section_name: entry.section_name || 'Umum',
                sectiondtl_name: entry.sectiondtl_name || '',
                checkpoint_name: entry.checkpoint_name || entry.checkPoint || '',
                result: entry.result || entry.hasilPenilaian || '',
                checked_by: entry.check_user_name || entry.checked_by_name || entry.checked_by || entry.check_user || entry.section_user_name || entry.section_user || entry.created_user || '-',
                section_date: entry.section_date || entry.doc_date || entry.created_date || '-',
                solution: entry.solution || entry.solusi || '',
                img_path: entry.img_path || entry.hasilFoto || entry.hasil_foto || entry.img || entry.foto || entry.image || entry.image_path || entry.path_foto || ''
            });
            return;
        }

        // Jika berupa nested Toyota format
        const areaName = entry.name || 'Showroom Toyota';
        const sections = entry.sections || [];

        sections.forEach(sec => {
            const secName = sec.name || 'Umum';
            const secNo = sec.no || '';
            const items = sec.items || [];

            items.forEach(it => {
                const imgRef = it.img_path || it.hasilFoto || it.hasil_foto || it.img || it.foto || it.image || it.image_path || it.path_foto || '';
                const checkUserVal = it.check_user_name || it.checked_by_name || it.checked_by || it.check_user || it.section_user_name || it.section_user || it.created_user || entry.check_user_name || entry.check_user || '-';
                const checkDateVal = it.section_date || it.doc_date || it.created_date || entry.doc_date || '-';

                itemsArray.push({
                    area_name: areaName,
                    section_name: secName,
                    section_no: secNo,
                    sectiondtl_name: '',
                    checkpoint_name: it.checkPoint || it.checkpoint_name || '',
                    result: it.hasilPenilaian || it.result || '',
                    checked_by: checkUserVal,
                    section_date: checkDateVal,
                    solution: it.solution || it.solusi || '',
                    img_path: imgRef
                });
            });
        });
    });

    return itemsArray;
}

/**
 * Generator Laporan Checkpoint TOYOTA (Sesuai Layout Presisi 9 Kolom)
 * - 9 Kolom Utama: No, Area, Section, Check Points, Hasil Penilaian, Checked By, Waktu Cek, Hasil Foto, Solusion.
 * - Legend keterangan (O = Standard, Δ = Tidak standard, X = Tidak tersedia) di sisi kanan.
 */
async function generateToyotaExcel(rawPayload, areaFilter = null, sectionFilter = null, subdetailFilter = null) {
    const items = normalizeToyotaData(rawPayload);
    let filteredItems = items;

    // Filter per Area jika ditentukan dari UI/API
    if (areaFilter && areaFilter !== 'ALL' && areaFilter !== 'Semua Area') {
        filteredItems = filteredItems.filter(item => {
            const area = String(item.area_name || '').trim().toLowerCase();
            return area === String(areaFilter).trim().toLowerCase();
        });
    }

    // Filter per Section jika ditentukan
    if (sectionFilter && sectionFilter !== 'ALL' && sectionFilter !== 'Semua Section') {
        filteredItems = filteredItems.filter(item => {
            const sec = String(item.section_name || '').trim().toLowerCase();
            return sec === String(sectionFilter).trim().toLowerCase();
        });
    }

    if (!filteredItems || filteredItems.length === 0) {
        filteredItems = items; // Fallback
    }

    const workbook = new ExcelJS.Workbook();
    const sheetTitle = 'Laporan Checkpoint Toyota';

    const worksheet = workbook.addWorksheet(sheetTitle, {
        views: [{ showGridLines: true }]
    });

    const FONT_FAMILY = 'Segoe UI';

    // Styles Sesuai Screenshot
    const fillHeaderGray = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };  // Light Gray Header
    const fillHeaderGreen = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } }; // Light Green Solusion Header
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

    // --- 1. BANNER JUDUL (Merged A1:I1) ---
    worksheet.mergeCells('A1:I1');
    const titleCell = worksheet.getCell('A1');
    
    let filterSubtitle = [];
    if (areaFilter && areaFilter !== 'ALL') filterSubtitle.push(`Area: ${areaFilter}`);
    if (sectionFilter && sectionFilter !== 'ALL') filterSubtitle.push(`Sec: ${sectionFilter}`);
    
    const bannerText = filterSubtitle.length > 0
        ? `  List Kondisi Fasilitas Cabang - TOYOTA [${filterSubtitle.join(' | ')}]`
        : '  List Kondisi Fasilitas Cabang - TOYOTA';

    titleCell.value = bannerText;
    titleCell.font = { name: FONT_FAMILY, size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = fillTitle;
    titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.getRow(1).height = 36;

    // --- 2. HEADER RINGKASAN METADATA DOKUMEN ---
    const firstItem = (filteredItems && filteredItems.length > 0) ? filteredItems[0] : {};
    const docNum = firstItem.doc_num || 'TYT/' + new Date().getFullYear() + '/001';
    const docDate = firstItem.doc_date || new Date().toISOString().split('T')[0];
    const checkUser = firstItem.check_user_name || firstItem.checked_by_name || firstItem.checked_by || firstItem.check_user || firstItem.section_user_name || firstItem.section_user || 'Auditor Toyota';

    const metaInfo = [
        ['No. Dokumen', docNum, 'Product Group', 'TOYOTA'],
        ['Tanggal Dokumen', docDate, 'Petugas Pemeriksa', checkUser],
        ['Status Laporan', 'SELESAI', 'Total Checkpoint', `${filteredItems.length} Item`]
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

    // --- 3. HEADER TABEL DATA (9 KOLOM: No, Area, Section, Check Points, Hasil Penilaian, Checked By, Waktu Cek, Hasil Foto, Solusion) ---
    const headers = [
        { header: 'No', key: 'no', width: 6 },
        { header: 'Area', key: 'area', width: 22 },
        { header: 'Section', key: 'section', width: 22 },
        { header: 'Check Points', key: 'question', width: 44 },
        { header: 'Hasil Penilaian', key: 'result', width: 14 },
        { header: 'Checked By', key: 'checked_by', width: 22 },
        { header: 'Waktu Cek', key: 'date', width: 20 },
        { header: 'Hasil Foto', key: 'photo', width: 75 },
        { header: 'Solusion', key: 'solution', width: 34 }
    ];

    worksheet.getRow(currentRow).height = 26;
    headers.forEach((h, colIdx) => {
        const colNum = colIdx + 1;
        const cell = worksheet.getCell(currentRow, colNum);
        cell.value = h.header;
        cell.alignment = alignCenter;
        cell.border = borderCell;
        worksheet.getColumn(colNum).width = h.width;

        if (h.key === 'solution') {
            cell.fill = fillHeaderGreen;
            cell.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF006100' } };
        } else {
            cell.fill = fillHeaderGray;
            cell.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FF1E293B' } };
        }
    });

    // LEGEND KETERANGAN DI SISI KANAN (KOLOM K)
    const legRow1 = worksheet.getCell(`K${currentRow}`);
    legRow1.value = 'O  = Standard';
    legRow1.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: 'FF1E293B' } };

    const legRow2 = worksheet.getCell(`K${currentRow + 1}`);
    legRow2.value = 'Δ  = Tidak standard';
    legRow2.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: 'FF1E293B' } };

    const legRow3 = worksheet.getCell(`K${currentRow + 2}`);
    legRow3.value = 'X  = Tidak tersedia';
    legRow3.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: 'FF1E293B' } };

    const headerRowIdx = currentRow;
    let rowCounter = 1;
    let maxFotoColWidth = 75;

    // --- 4. DATA BARIS MURNI ---
    for (let idx = 0; idx < filteredItems.length; idx++) {
        currentRow++;
        const item = filteredItems[idx];

        const areaName = item.area_name || 'Showroom Toyota';
        const sectionName = item.section_name || 'Section Umum';
        const cpName = item.checkpoint_name || '';
        const resultVal = String(item.result || '').trim();
        const checkedByVal = item.checked_by || item.check_user_name || item.check_user || item.section_user || '-';
        const secDate = item.section_date || item.doc_date || '-';
        const solutionText = String(item.solution || '').trim();
        const imgPathStr = item.img_path || item.hasilFoto || item.hasil_foto || item.img || item.foto || '';

        const isEven = (rowCounter % 2 === 0);

        // Cell Values (9 Kolom: 1=No, 2=Area, 3=Section, 4=Check Points, 5=Hasil Penilaian, 6=Checked By, 7=Waktu Cek, 8=Hasil Foto, 9=Solusion)
        const c1 = worksheet.getCell(currentRow, 1); c1.value = rowCounter; c1.alignment = alignCenter;
        const c2 = worksheet.getCell(currentRow, 2); c2.value = areaName; c2.alignment = alignCenter;
        const c3 = worksheet.getCell(currentRow, 3); c3.value = sectionName; c3.alignment = alignCenter;
        const c4 = worksheet.getCell(currentRow, 4); c4.value = cpName; c4.alignment = alignLeft;
        const c5 = worksheet.getCell(currentRow, 5); c5.value = resultVal; c5.alignment = alignCenter;
        const c6 = worksheet.getCell(currentRow, 6); c6.value = checkedByVal; c6.alignment = alignCenter;
        const c7 = worksheet.getCell(currentRow, 7); c7.value = secDate; c7.alignment = alignCenter;
        const c8 = worksheet.getCell(currentRow, 8); c8.alignment = alignCenter;
        const c9 = worksheet.getCell(currentRow, 9); c9.value = solutionText || '-'; c9.alignment = alignLeft;

        // Color Result Badge Toyota:
        const resUpper = resultVal.toUpperCase();
        if (resUpper === 'O' || resUpper === 'Y') {
            c5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
            c5.font = { name: FONT_FAMILY, size: 10.5, bold: true, color: { argb: 'FF15803D' } };
        } else if (resUpper === 'X' || resUpper === 'N') {
            c5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
            c5.font = { name: FONT_FAMILY, size: 10.5, bold: true, color: { argb: 'FFB91C1C' } };
        } else if (resUpper === '∆' || resUpper === 'DELTA') {
            c5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
            c5.font = { name: FONT_FAMILY, size: 10.5, bold: true, color: { argb: 'FFB45309' } };
        } else {
            c5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
            c5.font = { name: FONT_FAMILY, size: 9.5, bold: true, color: { argb: 'FF0369A1' } };
        }

        // Apply borders & font default
        for (let c = 1; c <= 9; c++) {
            const cell = worksheet.getCell(currentRow, c);
            cell.border = borderCell;
            if (c !== 5) { // Kecuali kolom penilaian (E)
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

            // Lebar kolom 8 (H) disesuaikan presisi dengan dimensi foto 220x150 px
            const neededColW = Math.round(imgData.widthPx / 7) + 6;
            if (neededColW > maxFotoColWidth) {
                maxFotoColWidth = neededColW;
            }
            worksheet.getColumn(8).width = maxFotoColWidth;

            // Tinggi baris disesuaikan presisi dengan tinggi foto 150px
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

    // AUTOFILTER EXCEL AKTIF PADA RANGE A7:I[lastRow]
    worksheet.autoFilter = {
        from: { row: headerRowIdx, column: 1 },
        to: { row: currentRow, column: 9 }
    };

    return await workbook.xlsx.writeBuffer();
}

module.exports = {
    generateToyotaExcel,
    normalizeToyotaData
};
