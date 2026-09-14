const ExcelJS = require('exceljs');
const { processCheckpointPhotos } = require('../imageHandler');

/**
 * Normalizer untuk mengubah data JSON Toyota (nested area -> sections -> items)
 * menjadi flat array item yang rapi.
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
                section_date: entry.section_date || entry.check_date || entry.doc_date || entry.created_date || '-',
                solution: entry.solution || entry.solusi || '',
                img_path: entry.img_path || entry.hasilFoto || entry.hasil_foto || entry.img || entry.foto || entry.image || entry.image_path || entry.path_foto || ''
            });
            return;
        }

        // Jika berupa nested Toyota format
        const areaName = entry.name || 'Showroom Toyota';
        const sections = entry.sections || [];

        sections.forEach((sec, secIdx) => {
            const secName = sec.name || 'Umum';
            const secNo = sec.no || (secIdx + 1);
            const items = sec.items || [];

            items.forEach(it => {
                const imgRef = it.img_path || it.hasilFoto || it.hasil_foto || it.img || it.foto || it.image || it.image_path || it.path_foto || '';
                const checkUserVal = it.check_user_name || it.checked_by_name || it.checked_by || it.check_user || it.section_user_name || it.section_user || it.created_user || entry.check_user_name || entry.check_user || '-';
                const checkDateVal = it.section_date || it.check_date || it.doc_date || it.created_date || entry.doc_date || '-';

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
 * Generator Laporan Checkpoint TOYOTA (Sesuai Screenshot Presisi - Section Block Merged Format)
 * - 7 Kolom Utama: No, Area, Section, Check Points, Hasil Penilaian, Hasil Foto, Solution.
 * - Kolom No, Area, Section, dan Hasil Foto di-merge secara vertikal per Section block.
 * - Pemrosesan Gambar Paralel Super Cepat (Promise.all).
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

    // Styles & Colors Sesuai Screenshot Toyota
    const fillHeaderGray = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };  // Light Gray Header
    const fillHeaderGreen = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } }; // Light Green Solution Header
    const fillTitle = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };        // Slate Dark Title

    const borderCell = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };

    const alignCenter = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const alignLeft = { horizontal: 'left', vertical: 'middle', wrapText: true };

    // --- 1. BANNER JUDUL (Merged A1:G1) ---
    worksheet.mergeCells('A1:G1');
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

    // --- 3. HEADER TABEL DATA (7 KOLOM SESUAI SCREENSHOT: No, Area, Section, Check Points, Hasil Penilaian, Hasil Foto, Solution) ---
    const headers = [
        { header: 'No', key: 'no', width: 6 },
        { header: 'Area', key: 'area', width: 24 },
        { header: 'Section', key: 'section', width: 22 },
        { header: 'Check Points', key: 'question', width: 44 },
        { header: 'Hasil Penilaian', key: 'result', width: 14 },
        { header: 'Hasil Foto', key: 'photo', width: 75 },
        { header: 'Solution', key: 'solution', width: 38 }
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

    // LEGEND KETERANGAN DI SISI KANAN (KOLOM I)
    const legRow1 = worksheet.getCell(`I${currentRow}`);
    legRow1.value = 'O  = Standard';
    legRow1.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: 'FF1E293B' } };

    const legRow2 = worksheet.getCell(`I${currentRow + 1}`);
    legRow2.value = 'Δ  = Tidak standard';
    legRow2.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: 'FF1E293B' } };

    const legRow3 = worksheet.getCell(`I${currentRow + 2}`);
    legRow3.value = 'X  = Tidak tersedia';
    legRow3.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: 'FF1E293B' } };

    const headerRowIdx = currentRow;

    // Group items by Section
    const sectionGroups = [];
    let currentGroup = null;

    filteredItems.forEach(item => {
        const secKey = `${item.area_name}___${item.section_name}`;
        if (!currentGroup || currentGroup.key !== secKey) {
            currentGroup = {
                key: secKey,
                area_name: item.area_name || 'Showroom Toyota',
                section_name: item.section_name || 'Section Umum',
                section_no: item.section_no || (sectionGroups.length + 1),
                items: []
            };
            sectionGroups.push(currentGroup);
        }
        currentGroup.items.push(item);
    });

    let maxFotoColWidth = 75;

    // PROSES SEMUA FOTO SECTION SECARA PARALEL (SUPER FAST!)
    const sectionImages = await Promise.all(
        sectionGroups.map(group => {
            const secPhotosStr = group.items
                .map(it => it.img_path || it.hasilFoto || it.hasil_foto || it.img || it.foto)
                .filter(Boolean)
                .join(',');
            return processCheckpointPhotos(secPhotosStr);
        })
    );

    // --- 4. RENDER BARIS SECTION BLOCK MERGED SESUAI SCREENSHOT TOYOTA ---
    for (let gIdx = 0; gIdx < sectionGroups.length; gIdx++) {
        const group = sectionGroups[gIdx];
        const numItems = group.items.length;
        const startRow = currentRow + 1;
        const endRow = startRow + numItems - 1;
        const imgData = sectionImages[gIdx];

        // Render tiap item baris pada Section ini
        for (let i = 0; i < numItems; i++) {
            const r = startRow + i;
            const item = group.items[i];

            const cpName = item.checkpoint_name || '';
            const resultVal = String(item.result || '').trim();
            const solutionText = String(item.solution || '').trim();

            const c4 = worksheet.getCell(r, 4); c4.value = cpName; c4.alignment = alignLeft;
            const c5 = worksheet.getCell(r, 5); c5.value = resultVal; c5.alignment = alignCenter;
            const c7 = worksheet.getCell(r, 7); c7.value = solutionText || '-'; c7.alignment = alignLeft;

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

            // Apply default border & font untuk baris r
            for (let c = 1; c <= 7; c++) {
                const cell = worksheet.getCell(r, c);
                cell.border = borderCell;
                if (c !== 5) {
                    cell.font = { name: FONT_FAMILY, size: 9.5, color: { argb: 'FF334155' } };
                }
            }
        }

        // --- MERGE CELLS SESUAI SCREENSHOT TOYOTA ---
        // 1. Merge Col A (No)
        if (numItems > 1) worksheet.mergeCells(startRow, 1, endRow, 1);
        const cNo = worksheet.getCell(startRow, 1);
        cNo.value = group.section_no;
        cNo.font = { name: FONT_FAMILY, size: 12, bold: true, color: { argb: 'FF0F172A' } };
        cNo.alignment = alignCenter;

        // 2. Merge Col B (Area)
        if (numItems > 1) worksheet.mergeCells(startRow, 2, endRow, 2);
        const cArea = worksheet.getCell(startRow, 2);
        cArea.value = group.area_name;
        cArea.font = { name: FONT_FAMILY, size: 11, bold: true, color: { argb: 'FF0F172A' } };
        cArea.alignment = alignCenter;

        // 3. Merge Col C (Section)
        if (numItems > 1) worksheet.mergeCells(startRow, 3, endRow, 3);
        const cSec = worksheet.getCell(startRow, 3);
        cSec.value = group.section_name;
        cSec.font = { name: FONT_FAMILY, size: 11, bold: true, color: { argb: 'FF0F172A' } };
        cSec.alignment = alignCenter;

        // 4. Merge Col F (Hasil Foto)
        if (numItems > 1) worksheet.mergeCells(startRow, 6, endRow, 6);
        const cPhoto = worksheet.getCell(startRow, 6);
        cPhoto.alignment = alignCenter;

        if (imgData) {
            const imageId = workbook.addImage({
                buffer: imgData.buffer,
                extension: 'png'
            });

            const neededColW = Math.round(imgData.widthPx / 7) + 6;
            if (neededColW > maxFotoColWidth) {
                maxFotoColWidth = neededColW;
            }
            worksheet.getColumn(6).width = maxFotoColWidth;

            // Hitung tinggi total section block agar foto muat sempurna
            const reqTotalH = Math.max(numItems * 28, Math.round((imgData.heightPx + 24) * 0.75));
            const rowH = Math.max(28, Math.round(reqTotalH / numItems));

            for (let r = startRow; r <= endRow; r++) {
                worksheet.getRow(r).height = rowH;
            }

            const rZero = startRow - 1;
            worksheet.addImage(imageId, {
                tl: { col: 5.04, row: rZero + 0.04 },
                ext: { width: imgData.widthPx, height: imgData.heightPx },
                editAs: 'oneCell'
            });
        } else {
            for (let r = startRow; r <= endRow; r++) {
                worksheet.getRow(r).height = 28;
            }
            cPhoto.value = '-';
        }

        // Terapkan border ke seluruh sel dalam blok section agar garis tabel utuh
        for (let r = startRow; r <= endRow; r++) {
            for (let c = 1; c <= 7; c++) {
                worksheet.getCell(r, c).border = borderCell;
            }
        }

        currentRow = endRow;
    }

    // AUTOFILTER EXCEL AKTIF PADA RANGE A7:G[lastRow]
    worksheet.autoFilter = {
        from: { row: headerRowIdx, column: 1 },
        to: { row: currentRow, column: 7 }
    };

    return await workbook.xlsx.writeBuffer();
}

module.exports = {
    generateToyotaExcel,
    normalizeToyotaData
};
