const fs = require('fs');
const path = require('path');
const axios = require('axios');
const sharp = require('sharp');

const THUMB_WIDTH = 120;
const THUMB_HEIGHT = 90;
const GAP = 10;
const PADDING = 6;
const BASE_IMAGE_URL = 'https://hrms.hasjrat.co.id/horor/';

/**
 * Membuat buffer foto placeholder jika foto asli tidak ditemukan / error.
 */
async function createPlaceholderImage(text = "Foto Tidak Ada", width = THUMB_WIDTH, height = THUMB_HEIGHT) {
    const svgText = `
    <svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
        <rect width="100%" height="100%" fill="#F0F3F6"/>
        <rect x="1" y="1" width="${width - 2}" height="${height - 2}" fill="none" stroke="#D0D5DF" stroke-width="1"/>
        <text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="11" fill="#828C9B">
            ${text}
        </text>
    </svg>`;
    return sharp(Buffer.from(svgText)).png().toBuffer();
}

/**
 * Membaca buffer foto dari URL (HTTP) atau File Lokal.
 * Jika path berupa relatif (seperti "hororupload/section/..."), otomatis ditambahkan BASE_IMAGE_URL.
 */
async function loadImageBuffer(imgRef) {
    if (!imgRef || typeof imgRef !== 'string') {
        return createPlaceholderImage("Path Kosong");
    }
    let cleanRef = imgRef.trim();
    if (!cleanRef) return createPlaceholderImage("Path Kosong");

    // Jika path belum diawali http:// atau https://
    if (!cleanRef.startsWith('http://') && !cleanRef.startsWith('https://')) {
        // Jika file lokal ada di disk, prioritaskan disk lokal
        if (fs.existsSync(cleanRef)) {
            return fs.readFileSync(cleanRef);
        }
        const baseName = path.basename(cleanRef);
        if (fs.existsSync(baseName)) {
            return fs.readFileSync(baseName);
        }

        // Otomatis tambahkan Prefix Domain URL: https://hrms.hasjrat.co.id/horor/
        if (cleanRef.startsWith('/')) {
            cleanRef = 'https://hrms.hasjrat.co.id/horor' + cleanRef;
        } else {
            cleanRef = BASE_IMAGE_URL + cleanRef;
        }
    }

    try {
        const response = await axios.get(cleanRef, {
            responseType: 'arraybuffer',
            timeout: 10000,
            headers: {
                'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)'
            }
        });
        return Buffer.from(response.data);
    } catch (err) {
        console.warn(`Gagal mendownload gambar dari ${cleanRef}: ${err.message}`);
        return createPlaceholderImage("Err Download");
    }
}

/**
 * Meresize foto ke 120x90px dengan crop center dan border.
 */
async function processSinglePhoto(imgRef) {
    const rawBuffer = await loadImageBuffer(imgRef);
    try {
        const borderOverlay = Buffer.from(
            `<svg width="${THUMB_WIDTH}" height="${THUMB_HEIGHT}"><rect x="0" y="0" width="${THUMB_WIDTH}" height="${THUMB_HEIGHT}" fill="none" stroke="#DAE0E9" stroke-width="1"/></svg>`
        );

        return await sharp(rawBuffer)
            .resize(THUMB_WIDTH, THUMB_HEIGHT, { fit: 'cover', position: 'center' })
            .composite([{ input: borderOverlay }])
            .png()
            .toBuffer();
    } catch (e) {
        return createPlaceholderImage("Err Format");
    }
}

/**
 * Memproses string `img_path` (bisa multi foto dipisah koma).
 * Mengembalikan { buffer, widthPx, heightPx, count }
 */
async function processCheckpointPhotos(imgPathStr) {
    if (!imgPathStr || typeof imgPathStr !== 'string') {
        return null;
    }

    const paths = imgPathStr.split(',').map(p => p.trim()).filter(Boolean);
    if (paths.length === 0) {
        return null;
    }

    const thumbBuffers = await Promise.all(paths.map(p => processSinglePhoto(p)));
    const count = thumbBuffers.length;

    const totalWidth = (PADDING * 2) + (count * THUMB_WIDTH) + ((count - 1) * GAP);
    const totalHeight = (PADDING * 2) + THUMB_HEIGHT;

    // Outer border & background canvas SVG
    const bgSvg = Buffer.from(
        `<svg width="${totalWidth}" height="${totalHeight}">
            <rect width="100%" height="100%" fill="#F8FAFC"/>
            <rect x="0" y="0" width="${totalWidth}" height="${totalHeight}" fill="none" stroke="#E2E8F0" stroke-width="1"/>
        </svg>`
    );

    const compositeLayers = [
        { input: bgSvg, top: 0, left: 0 }
    ];

    let xOffset = PADDING;
    for (const thumbBuf of thumbBuffers) {
        compositeLayers.push({
            input: thumbBuf,
            top: PADDING,
            left: xOffset
        });
        xOffset += THUMB_WIDTH + GAP;
    }

    const finalBuffer = await sharp({
        create: {
            width: totalWidth,
            height: totalHeight,
            channels: 4,
            background: { r: 248, g: 250, b: 252, alpha: 1 }
        }
    })
    .composite(compositeLayers)
    .png()
    .toBuffer();

    return {
        buffer: finalBuffer,
        widthPx: totalWidth,
        heightPx: totalHeight,
        count: count
    };
}

module.exports = {
    processCheckpointPhotos,
    BASE_IMAGE_URL
};
