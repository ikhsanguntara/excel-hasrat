const fs = require('fs');
const path = require('path');
const https = require('https');
const axios = require('axios');
const sharp = require('sharp');

// Ukuran Thumbnail Foto yang Besar & Wide (220x150 px - Rasio HD)
const THUMB_WIDTH = 220;
const THUMB_HEIGHT = 150;
const GAP = 12;
const PADDING = 8;
const BASE_IMAGE_URL = 'https://hrms.hasjrat.co.id/horor/';

// Https Agent dengan rejectUnauthorized: false untuk menangani SSL sertifikat server Hasjrat
const httpsAgent = new https.Agent({
    rejectUnauthorized: false
});

/**
 * Membuat buffer foto placeholder jika foto asli tidak ditemukan / error.
 */
async function createPlaceholderImage(text = "Foto Kosong", width = THUMB_WIDTH, height = THUMB_HEIGHT) {
    const svgText = `
    <svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
        <rect width="100%" height="100%" fill="#F0F3F6"/>
        <rect x="1" y="1" width="${width - 2}" height="${height - 2}" fill="none" stroke="#D0D5DF" stroke-width="1.5"/>
        <text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="13" font-weight="500" fill="#828C9B">
            ${text}
        </text>
    </svg>`;
    return sharp(Buffer.from(svgText)).png().toBuffer();
}

/**
 * Membaca buffer foto dari URL (HTTP/HTTPS) atau File Lokal.
 */
async function loadImageBuffer(imgRef) {
    if (!imgRef || typeof imgRef !== 'string') {
        return createPlaceholderImage("Foto Kosong");
    }
    let cleanRef = String(imgRef).trim().replace(/\\/g, '/');
    if (!cleanRef) return createPlaceholderImage("Foto Kosong");

    // Jika path belum diawali http:// atau https://
    if (!cleanRef.startsWith('http://') && !cleanRef.startsWith('https://')) {
        // Jika file lokal ada di disk lokal, prioritaskan disk lokal
        if (fs.existsSync(cleanRef)) {
            return fs.readFileSync(cleanRef);
        }
        const baseName = path.basename(cleanRef);
        if (fs.existsSync(baseName)) {
            return fs.readFileSync(baseName);
        }

        // Hapus slash di paling depan jika ada
        if (cleanRef.startsWith('/')) {
            cleanRef = cleanRef.substring(1);
        }

        // Mencegah duplikasi URL "horor/horor/hororupload..."
        if (cleanRef.startsWith('horor/')) {
            cleanRef = cleanRef.substring(6);
        }

        // Konstruksi URL full
        cleanRef = BASE_IMAGE_URL + cleanRef;
    }

    try {
        const encodedUrl = encodeURI(cleanRef);
        const response = await axios.get(encodedUrl, {
            responseType: 'arraybuffer',
            timeout: 15000,
            httpsAgent: httpsAgent,
            headers: {
                'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
        });
        return Buffer.from(response.data);
    } catch (err) {
        console.warn(`Gagal mendownload gambar dari ${cleanRef}: ${err.message}`);
        return createPlaceholderImage("Foto Lampiran");
    }
}

/**
 * Meresize foto ke ukuran seragam yang besar & lebar (220x150px) dengan crop center & border presisi.
 */
async function processSinglePhoto(imgRef) {
    const rawBuffer = await loadImageBuffer(imgRef);
    try {
        const borderOverlay = Buffer.from(
            `<svg width="${THUMB_WIDTH}" height="${THUMB_HEIGHT}"><rect x="0" y="0" width="${THUMB_WIDTH}" height="${THUMB_HEIGHT}" fill="none" stroke="#CBD5E1" stroke-width="1.5"/></svg>`
        );

        return await sharp(rawBuffer)
            .resize(THUMB_WIDTH, THUMB_HEIGHT, { 
                fit: 'cover', 
                position: 'center',
                withoutEnlargement: false 
            })
            .composite([{ input: borderOverlay }])
            .png()
            .toBuffer();
    } catch (e) {
        return createPlaceholderImage("Format Gambar");
    }
}

/**
 * Memproses string path foto (bisa multi foto dipisah koma `,`, titik koma `;`, atau baris baru).
 * Mengembalikan { buffer, widthPx, heightPx, count }
 */
async function processCheckpointPhotos(imgPathStr) {
    if (!imgPathStr || typeof imgPathStr !== 'string') {
        return null;
    }

    // Split berdasarkan koma, titik koma, atau baris baru
    const paths = imgPathStr
        .split(/[,;\n]+/)
        .map(p => p.trim())
        .filter(p => p.length > 0);

    if (paths.length === 0) {
        return null;
    }

    // Process & resize semua foto secara paralel ke dimensi besar 220x150px
    const thumbBuffers = await Promise.all(paths.map(p => processSinglePhoto(p)));
    const count = thumbBuffers.length;

    const totalWidth = (PADDING * 2) + (count * THUMB_WIDTH) + ((count - 1) * GAP);
    const totalHeight = (PADDING * 2) + THUMB_HEIGHT;

    // Canvas latar belakang bersih dengan border luar halus
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
    BASE_IMAGE_URL,
    THUMB_WIDTH,
    THUMB_HEIGHT
};
