import os
import io
import requests
from PIL import Image, ImageDraw, ImageFont

THUMB_WIDTH = 120
THUMB_HEIGHT = 90
GAP = 10
PADDING = 6
BG_COLOR = (248, 249, 250)
BORDER_COLOR = (218, 224, 233)
BASE_IMAGE_URL = "https://hrms.hasjrat.co.id/horor/"

def create_placeholder_image(text="Foto Tidak Tersedia", width=THUMB_WIDTH, height=THUMB_HEIGHT) -> Image.Image:
    img = Image.new("RGB", (width, height), color=(240, 243, 246))
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, width - 1, height - 1], outline=(200, 205, 215), width=1)
    try:
        font = ImageFont.load_default()
    except Exception:
        font = None
    draw.text((width // 2, height // 2), text, fill=(130, 140, 155), anchor="mm", font=font)
    return img

def load_single_image(img_ref: str) -> Image.Image:
    img_ref = img_ref.strip()
    if not img_ref:
        return create_placeholder_image("Path Kosong")

    # Jika belum diawali http:// atau https://
    if not img_ref.startswith("http://") and not img_ref.startswith("https://"):
        if os.path.exists(img_ref):
            return Image.open(img_ref).convert("RGB")
        basename = os.path.basename(img_ref)
        if os.path.exists(basename):
            return Image.open(basename).convert("RGB")
        
        # Tambahkan base URL otomatis
        if img_ref.startswith("/"):
            img_ref = "https://hrms.hasjrat.co.id/horor" + img_ref
        else:
            img_ref = BASE_IMAGE_URL + img_ref

    try:
        resp = requests.get(img_ref, timeout=10, headers={'User-Agent': 'Mozilla/5.0'})
        if resp.status_code == 200:
            return Image.open(io.BytesIO(resp.content)).convert("RGB")
        else:
            return create_placeholder_image(f"HTTP {resp.status_code}")
    except Exception as e:
        return create_placeholder_image("Err Download")

def resize_and_crop(img: Image.Image, target_w=THUMB_WIDTH, target_h=THUMB_HEIGHT) -> Image.Image:
    orig_w, orig_h = img.size
    if orig_w == 0 or orig_h == 0:
        return create_placeholder_image("Err Dim", target_w, target_h)
    
    aspect = orig_w / orig_h
    target_aspect = target_w / target_h
    
    if aspect > target_aspect:
        new_h = target_h
        new_w = int(target_h * aspect)
    else:
        new_w = target_w
        new_h = int(target_w / aspect)
        
    resizing_filter = getattr(Image, 'Resampling', Image).LANCZOS
    img_resized = img.resize((new_w, new_h), resizing_filter)
    
    left = (new_w - target_w) // 2
    top = (new_h - target_h) // 2
    right = left + target_w
    bottom = top + target_h
    
    cropped = img_resized.crop((left, top, right, bottom))
    draw = ImageDraw.Draw(cropped)
    draw.rectangle([0, 0, target_w - 1, target_h - 1], outline=BORDER_COLOR, width=1)
    return cropped

def process_checkpoint_photos(img_path_str: str) -> tuple[io.BytesIO, float, float]:
    if not img_path_str or not isinstance(img_path_str, str):
        paths = []
    else:
        paths = [p.strip() for p in img_path_str.split(",") if p.strip()]
    
    if not paths:
        return None, 0, 0
    
    processed_thumbs = []
    for p in paths:
        raw_img = load_single_image(p)
        thumb = resize_and_crop(raw_img, THUMB_WIDTH, THUMB_HEIGHT)
        processed_thumbs.append(thumb)
    
    count = len(processed_thumbs)
    total_w = PADDING * 2 + (count * THUMB_WIDTH) + ((count - 1) * GAP)
    total_h = PADDING * 2 + THUMB_HEIGHT
    
    canvas = Image.new("RGB", (total_w, total_h), color=BG_COLOR)
    draw = ImageDraw.Draw(canvas)
    draw.rectangle([0, 0, total_w - 1, total_h - 1], outline=(230, 235, 242), width=1)
    
    x_offset = PADDING
    for thumb in processed_thumbs:
        canvas.paste(thumb, (x_offset, PADDING))
        x_offset += THUMB_WIDTH + GAP
    
    img_byte_arr = io.BytesIO()
    canvas.save(img_byte_arr, format='PNG')
    img_byte_arr.seek(0)
    
    return img_byte_arr, total_w, total_h
