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

// In-memory cache & In-flight request deduplicator (Mencegah request HTTP berulang ke server)
const rawImageCache = new Map();
const inFlightRequests = new Map();
const processedThumbCache = new Map();
const inFlightThumbs = new Map();

// Https Agent dengan rejectUnauthorized: false & keep-alive
const httpsAgent = new https.Agent({
    rejectUnauthorized: false,
    keepAlive: true,
    maxSockets: 50
});

// Axios instance
const apiClient = axios.create({
    httpsAgent: httpsAgent,
    timeout: 10000,
    responseType: 'arraybuffer',
    headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
});

/**
 * Membuat buffer foto placeholder jika foto asli tidak ditemukan / error.
 */
async function createPlaceholderImage(text = "Foto Kosong", width = THUMB_WIDTH, height = THUMB_HEIGHT) {
    const cacheKey = `placeholder_${text}_${width}_${height}`;
    if (processedThumbCache.has(cacheKey)) {
        return processedThumbCache.get(cacheKey);
    }

    const svgText = `
    <svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
        <rect width="100%" height="100%" fill="#F0F3F6"/>
        <rect x="1" y="1" width="${width - 2}" height="${height - 2}" fill="none" stroke="#D0D5DF" stroke-width="1.5"/>
        <text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="13" font-weight="500" fill="#828C9B">
            ${text}
        </text>
    </svg>`;
    const buf = await sharp(Buffer.from(svgText)).png().toBuffer();
    processedThumbCache.set(cacheKey, buf);
    return buf;
}

/**
 * Membaca buffer foto dari URL (HTTP/HTTPS) atau File Lokal dengan In-Memory Caching & In-Flight Deduplication.
 */
async function loadImageBuffer(imgRef) {
    if (!imgRef || typeof imgRef !== 'string') {
        return createPlaceholderImage("Foto Kosong");
    }
    let cleanRef = String(imgRef).trim().replace(/\\/g, '/');
    if (!cleanRef) return createPlaceholderImage("Foto Kosong");

    // Jika path belum diawali http:// atau https://
    if (!cleanRef.startsWith('http://') && !cleanRef.startsWith('https://')) {
        if (fs.existsSync(cleanRef)) {
            return fs.readFileSync(cleanRef);
        }
        const baseName = path.basename(cleanRef);
        if (fs.existsSync(baseName)) {
            return fs.readFileSync(baseName);
        }

        if (cleanRef.startsWith('/')) {
            cleanRef = cleanRef.substring(1);
        }

        if (cleanRef.startsWith('horor/')) {
            cleanRef = cleanRef.substring(6);
        }

        cleanRef = BASE_IMAGE_URL + cleanRef;
    }

    // 1. Cek Memory Cache (0ms Instant!)
    if (rawImageCache.has(cleanRef)) {
        return rawImageCache.get(cleanRef);
    }

    // 2. Cek apakah URL ini sedang dalam proses download (Deduplikasi In-Flight Request)
    if (inFlightRequests.has(cleanRef)) {
        return inFlightRequests.get(cleanRef);
    }

    const downloadPromise = (async () => {
        try {
            const encodedUrl = encodeURI(cleanRef);
            const response = await apiClient.get(encodedUrl);
            const buf = Buffer.from(response.data);
            rawImageCache.set(cleanRef, buf);
            return buf;
        } catch (err) {
            console.warn(`Gagal mendownload gambar dari ${cleanRef}: ${err.message}`);
            const placeholder = await createPlaceholderImage("Foto Kosong");
            rawImageCache.set(cleanRef, placeholder);
            return placeholder;
        } finally {
            inFlightRequests.delete(cleanRef);
        }
    })();

    inFlightRequests.set(cleanRef, downloadPromise);
    return downloadPromise;
}

/**
 * Meresize foto ke ukuran seragam yang besar & lebar (220x150px) dengan crop center & border presisi.
 */
async function processSinglePhoto(imgRef) {
    const cacheKey = `thumb_${imgRef}_${THUMB_WIDTH}_${THUMB_HEIGHT}`;
    if (processedThumbCache.has(cacheKey)) {
        return processedThumbCache.get(cacheKey);
    }
    if (inFlightThumbs.has(cacheKey)) {
        return inFlightThumbs.get(cacheKey);
    }

    const thumbPromise = (async () => {
        const rawBuffer = await loadImageBuffer(imgRef);
        try {
            const borderOverlay = Buffer.from(
                `<svg width="${THUMB_WIDTH}" height="${THUMB_HEIGHT}"><rect x="0" y="0" width="${THUMB_WIDTH}" height="${THUMB_HEIGHT}" fill="none" stroke="#CBD5E1" stroke-width="1.5"/></svg>`
            );

            const resized = await sharp(rawBuffer)
                .resize(THUMB_WIDTH, THUMB_HEIGHT, { 
                    fit: 'cover', 
                    position: 'center',
                    withoutEnlargement: false 
                })
                .composite([{ input: borderOverlay }])
                .png()
                .toBuffer();

            processedThumbCache.set(cacheKey, resized);
            return resized;
        } catch (e) {
            const placeholder = await createPlaceholderImage("Format Gambar");
            processedThumbCache.set(cacheKey, placeholder);
            return placeholder;
        } finally {
            inFlightThumbs.delete(cacheKey);
        }
    })();

    inFlightThumbs.set(cacheKey, thumbPromise);
    return thumbPromise;
}

/**
 * Memproses path/URL foto (bisa string koma/baris baru atau Array URL).
 * Menggabungkan hingga N foto menjadi composite grid horizontal rapi.
 * Mengembalikan { buffer, widthPx, heightPx, count }
 */
async function processCheckpointPhotos(imgPathInput) {
    if (!imgPathInput) {
        return null;
    }

    let rawPaths = [];
    if (Array.isArray(imgPathInput)) {
        rawPaths = imgPathInput
            .map(p => {
                if (typeof p === 'string') return p.trim();
                if (p && typeof p === 'object') return (p.url || p.img_path || p.foto || p.path || '').trim();
                return String(p || '').trim();
            })
            .filter(p => p.length > 0);
    } else if (typeof imgPathInput === 'string') {
        rawPaths = imgPathInput
            .split(/[,;\n]+/)
            .map(p => p.trim())
            .filter(p => p.length > 0);
    } else {
        return null;
    }

    const paths = [...new Set(rawPaths)];

    if (paths.length === 0) {
        return null;
    }

    const compositeCacheKey = `composite_${paths.join('___')}_${THUMB_WIDTH}_${THUMB_HEIGHT}`;
    if (processedThumbCache.has(compositeCacheKey)) {
        return processedThumbCache.get(compositeCacheKey);
    }

    const thumbBuffers = await Promise.all(paths.map(p => processSinglePhoto(p)));
    const count = thumbBuffers.length;

    const totalWidth = (PADDING * 2) + (count * THUMB_WIDTH) + ((count - 1) * GAP);
    const totalHeight = (PADDING * 2) + THUMB_HEIGHT;

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

    const result = {
        buffer: finalBuffer,
        widthPx: totalWidth,
        heightPx: totalHeight,
        count: count
    };

    processedThumbCache.set(compositeCacheKey, result);
    return result;
}

module.exports = {
    processCheckpointPhotos,
    BASE_IMAGE_URL,
    THUMB_WIDTH,
    THUMB_HEIGHT
};
