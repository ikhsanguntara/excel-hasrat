import io
from typing import List, Dict, Any
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenPyxlImage

from app.image_handler import process_checkpoint_photos

def generate_checkpoint_excel(items: List[Dict[str, Any]]) -> io.BytesIO:
    """
    Menerima list dictionary data checkpoint dan menghasilkan Excel workbook (.xlsx)
    dengan format yang sangat rapi, grouping per Area & Section, serta foto yang ter-resize seragam.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Laporan Checkpoint"
    
    # Pastikan gridlines terlihat
    ws.views.sheetView[0].showGridLines = True

    # --- DEFINISI PALET WARNA & STYLES ---
    FONT_FAMILY = "Segoe UI"
    
    # Fonts
    title_font = Font(name=FONT_FAMILY, size=16, bold=True, color="FFFFFF")
    subtitle_font = Font(name=FONT_FAMILY, size=10, italic=True, color="E2E8F0")
    header_font = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
    area_font = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    section_font = Font(name=FONT_FAMILY, size=10.5, bold=True, color="FFFFFF")
    subdetail_font = Font(name=FONT_FAMILY, size=10, bold=True, color="1E293B")
    data_font = Font(name=FONT_FAMILY, size=9.5, color="334155")
    meta_label_font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="475569")
    meta_val_font = Font(name=FONT_FAMILY, size=9.5, color="0F172A")

    # Fills (Backgrounds)
    fill_title = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")       # Dark Slate Navy
    fill_th = PatternFill(start_color="334155", end_color="334155", fill_type="solid")          # Header Columns
    fill_area = PatternFill(start_color="475569", end_color="475569", fill_type="solid")        # Area Header
    fill_section = PatternFill(start_color="64748B", end_color="64748B", fill_type="solid")     # Section Header
    fill_subdetail = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")   # Section Detail Header
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")       # Zebra row
    
    # Result Badge Fills & Fonts
    result_styles = {
        "Y": {
            "fill": PatternFill(start_color="DCFCE7", fill_type="solid"),
            "font": Font(name=FONT_FAMILY, size=10, bold=True, color="15803D")
        },
        "N": {
            "fill": PatternFill(start_color="FEE2E2", fill_type="solid"),
            "font": Font(name=FONT_FAMILY, size=10, bold=True, color="B91C1C")
        },
        "DEFAULT": {
            "fill": PatternFill(start_color="E0F2FE", fill_type="solid"),
            "font": Font(name=FONT_FAMILY, size=10, bold=True, color="0369A1")
        }
    }

    # Borders
    thin_side = Side(style='thin', color="CBD5E1")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_box_top = Border(top=Side(style='medium', color="94A3B8"))
    border_box_bottom = Border(bottom=Side(style='medium', color="94A3B8"))

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    # --- 1. BANNER JUDUL LAPORAN ---
    ws.merge_cells("A1:F1")
    cell_title = ws["A1"]
    cell_title.value = "  LAPORAN HASIL CHECKPOINT & INSPEKSI"
    cell_title.font = title_font
    cell_title.fill = fill_title
    cell_title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 40

    # --- 2. HEADER RINGKASAN METADATA DOKUMEN ---
    first_item = items[0] if items else {}
    doc_num = first_item.get("doc_num", "-")
    doc_date = first_item.get("doc_date", "-")
    check_user = first_item.get("check_user", "-")
    product_group = first_item.get("product_group", "-")
    start_date = first_item.get("start_doc_date", "-")
    end_date = first_item.get("end_doc_date", "-")

    # Baris Metadata (Row 3-5)
    meta_info = [
        ("No. Dokumen", doc_num, "Product Group", product_group),
        ("Tanggal Dokumen", doc_date, "Petugas Pemeriksa", check_user),
        ("Periode Inspeksi", f"{start_date} s/d {end_date}", "Total Checkpoint", f"{len(items)} Item")
    ]

    for idx, (l1, v1, l2, v2) in enumerate(meta_info, start=3):
        ws.row_dimensions[idx].height = 20
        # Col A: Label 1
        c_l1 = ws.cell(row=idx, column=1, value=l1)
        c_l1.font = meta_label_font
        c_l1.alignment = align_left
        
        # Col B: Val 1
        c_v1 = ws.cell(row=idx, column=2, value=v1)
        c_v1.font = meta_val_font
        c_v1.alignment = align_left
        
        # Col D: Label 2
        c_l2 = ws.cell(row=idx, column=4, value=l2)
        c_l2.font = meta_label_font
        c_l2.alignment = align_left
        
        # Col E: Val 2
        c_v2 = ws.cell(row=idx, column=5, value=v2)
        c_v2.font = meta_val_font
        c_v2.alignment = align_left

    current_row = 7

    # --- 3. HEADER TABEL DATA ---
    headers = [
        ("No", 6),
        ("Kode / Mandat", 15),
        ("Pertanyaan Checkpoint", 48),
        ("Hasil", 14),
        ("Waktu Cek", 22),
        ("Foto Lampiran", 55)
    ]

    ws.row_dimensions[current_row].height = 28
    for col_idx, (h_text, col_w) in enumerate(headers, start=1):
        cell = ws.cell(row=current_row, column=col_idx, value=h_text)
        cell.font = header_font
        cell.fill = fill_th
        cell.alignment = align_center
        cell.border = border_cell
        # Set initial column width
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_w

    current_row += 1

    # --- 4. DATA GROUPING & POPULASI ROW ---
    # Kelompokkan data per area_name -> section_name
    grouped_data: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}
    for item in items:
        area = item.get("area_name") or "Area Tidak Terdefinisi"
        sec = item.get("section_name") or "Section Umum"
        if area not in grouped_data:
            grouped_data[area] = {}
        if sec not in grouped_data[area]:
            grouped_data[area][sec] = []
        grouped_data[area][sec].append(item)

    row_counter = 1
    max_foto_col_width = 55  # Dinamis menyesuaikan foto terbanyak

    for area_name, sections in grouped_data.items():
        # --- HEADER AREA ---
        ws.row_dimensions[current_row].height = 24
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
        area_cell = ws.cell(row=current_row, column=1, value=f" AREA: {area_name.upper()}")
        area_cell.font = area_font
        area_cell.fill = fill_area
        area_cell.alignment = Alignment(horizontal="left", vertical="center")
        for col_i in range(1, 7):
            ws.cell(row=current_row, column=col_i).border = border_cell
        current_row += 1

        for section_name, sec_items in sections.items():
            # --- HEADER SECTION ---
            ws.row_dimensions[current_row].height = 22
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
            sec_cell = ws.cell(row=current_row, column=1, value=f"   Section: {section_name}")
            sec_cell.font = section_font
            sec_cell.fill = fill_section
            sec_cell.alignment = Alignment(horizontal="left", vertical="center")
            for col_i in range(1, 7):
                ws.cell(row=current_row, column=col_i).border = border_cell
            current_row += 1

            # Render item per section
            last_sec_dtl = None
            for item_idx, item in enumerate(sec_items):
                sec_dtl = item.get("sectiondtl_name")
                
                # Header Sub-Detail jika ada perubahan sub-detail
                if sec_dtl and sec_dtl != last_sec_dtl:
                    ws.row_dimensions[current_row].height = 20
                    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
                    dtl_cell = ws.cell(row=current_row, column=1, value=f"     • Sub-bagian: {sec_dtl}")
                    dtl_cell.font = subdetail_font
                    dtl_cell.fill = fill_subdetail
                    dtl_cell.alignment = Alignment(horizontal="left", vertical="center")
                    for col_i in range(1, 7):
                        ws.cell(row=current_row, column=col_i).border = border_cell
                    current_row += 1
                    last_sec_dtl = sec_dtl

                # --- POPULASI BARIS ITEM CHECKPOINT ---
                seq_num = item.get("seq_num") or row_counter
                is_mandatory = item.get("is_mandatory", "N")
                code_str = f"#{seq_num}" + (" (Wajib)" if is_mandatory == "Y" else "")
                cp_name = item.get("checkpoint_name", "")
                result_val = str(item.get("result", "")).strip()
                sec_date = item.get("section_date", "")
                img_path_str = item.get("img_path", "")

                # Cell 1: No
                c1 = ws.cell(row=current_row, column=1, value=row_counter)
                c1.alignment = align_center
                
                # Cell 2: Kode
                c2 = ws.cell(row=current_row, column=2, value=code_str)
                c2.alignment = align_center
                
                # Cell 3: Pertanyaan
                c3 = ws.cell(row=current_row, column=3, value=cp_name)
                c3.alignment = align_left
                
                # Cell 4: Hasil (Dengan Badge Warna)
                c4 = ws.cell(row=current_row, column=4, value=result_val)
                c4.alignment = align_center
                res_st = result_styles.get(result_val.upper(), result_styles["DEFAULT"])
                c4.fill = res_st["fill"]
                c4.font = res_st["font"]
                
                # Cell 5: Tanggal
                c5 = ws.cell(row=current_row, column=5, value=sec_date)
                c5.alignment = align_center
                
                # Cell 6: Foto
                c6 = ws.cell(row=current_row, column=6)
                c6.alignment = align_center

                # Apply font & border dasar ke seluruh sel di baris ini
                is_even = (row_counter % 2 == 0)
                for col_i in range(1, 7):
                    cell_i = ws.cell(row=current_row, column=col_i)
                    if col_i != 4:  # Kecuali kolom hasil yang punya warna sendiri
                        cell_i.font = data_font
                        if is_even:
                            cell_i.fill = fill_zebra
                    cell_i.border = border_cell

                # --- PROSES DAN SISIPKAN FOTO HASIL RESIZING ---
                img_stream, img_w, img_h = process_checkpoint_photos(img_path_str)
                
                if img_stream:
                    # Buat objek gambar openpyxl
                    img_obj = OpenPyxlImage(img_stream)
                    # Tentukan ukuran sel Excel (dalam points/pixels)
                    img_obj.width = img_w
                    img_obj.height = img_h
                    
                    # Tambahkan ke sel F (kolom 6)
                    col_letter = get_column_letter(6)
                    ws.add_image(img_obj, f"{col_letter}{current_row}")
                    
                    # Hitung height baris (1 pixel ~ 0.75 pt, berikan sedikit margin padding)
                    calc_row_height = max(85, int((img_h + 16) * 0.75))
                    ws.row_dimensions[current_row].height = calc_row_height
                    
                    # Sesuaikan lebar kolom F jika kanvas foto lebih lebar
                    needed_col_w = int(img_w / 7) + 4
                    if needed_col_w > max_foto_col_width:
                        max_foto_col_width = needed_col_w
                        ws.column_dimensions["F"].width = max_foto_col_width
                else:
                    ws.row_dimensions[current_row].height = 28
                    c6.value = "-"

                row_counter += 1
                current_row += 1

    # Save to BytesIO buffer
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream
