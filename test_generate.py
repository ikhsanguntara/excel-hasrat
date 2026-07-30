import json
import os
from app.excel_generator import generate_checkpoint_excel

def test_checkpoint_generation():
    json_path = "checkpoint_data.json"
    output_path = "laporan_checkpoint_sample.xlsx"
    
    if not os.path.exists(json_path):
        print(f"Error: {json_path} tidak ditemukan!")
        return

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    print(f"Membaca {len(data)} item dari {json_path}...")
    
    excel_stream = generate_checkpoint_excel(data)
    
    with open(output_path, "wb") as f_out:
        f_out.write(excel_stream.getvalue())

    print(f"SUCCESS! File Excel berhasil dibuat di: {os.path.abspath(output_path)}")
    print(f"Ukuran file: {os.path.getsize(output_path)} bytes")

if __name__ == "__main__":
    test_checkpoint_generation()
