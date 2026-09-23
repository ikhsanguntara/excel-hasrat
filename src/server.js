const os = require('os');
const path = require('path');
const express = require('express');
const { generateCheckpointExcel } = require('./excelGenerator');
const { getImageCacheStats } = require('./imageHandler');
const packageInfo = require('../package.json');

const app = express();
const PORT = process.env.PORT || 8000;

// Middleware for JSON body parsing (limit 50mb for large JSONs)
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Helper untuk format uptime
function formatUptime(seconds) {
    const d = Math.floor(seconds / (3600 * 24));
    const h = Math.floor((seconds % (3600 * 24)) / 3600);
    const m = Math.floor((seconds % 3600) / 60);
    const s = Math.floor(seconds % 60);
    const parts = [];
    if (d > 0) parts.push(`${d}d`);
    if (h > 0) parts.push(`${h}h`);
    if (m > 0) parts.push(`${m}m`);
    parts.push(`${s}s`);
    return parts.join(' ');
}

// Handler Health Check (Rich System, Service & Metrics Info)
const handleHealthCheck = (req, res) => {
    const uptimeSec = process.uptime();
    const mem = process.memoryUsage();
    const cpus = os.cpus();

    res.status(200).json({
        status: 'OK',
        message: 'Excel Generator Service is healthy and running',
        service: {
            name: packageInfo.name || 'excel_ai',
            version: packageInfo.version || '1.0.0',
            description: packageInfo.description || 'JSON to Excel Inspection Report Generator API'
        },
        timestamp: new Date().toISOString(),
        uptime: formatUptime(uptimeSec),
        uptime_seconds: Math.floor(uptimeSec),
        system: {
            hostname: os.hostname(),
            platform: `${os.platform()} (${os.type()})`,
            arch: os.arch(),
            release: os.release(),
            node_version: process.version,
            pid: process.pid,
            cpu_cores: cpus ? cpus.length : 0,
            cpu_model: cpus && cpus[0] ? cpus[0].model : 'Unknown',
            total_memory: `${(os.totalmem() / (1024 ** 3)).toFixed(2)} GB`,
            free_memory: `${(os.freemem() / (1024 ** 3)).toFixed(2)} GB`
        },
        process_memory: {
            rss: `${Math.round(mem.rss / 1024 / 1024)} MB`,
            heapTotal: `${Math.round(mem.heapTotal / 1024 / 1024)} MB`,
            heapUsed: `${Math.round(mem.heapUsed / 1024 / 1024)} MB`,
            external: `${Math.round(mem.external / 1024 / 1024)} MB`
        },
        features: {
            supported_reports: ['asset-counting', 'toyota', 'yamaha'],
            image_cache: getImageCacheStats(),
            excel_engine: 'ExcelJS v4.4.0',
            image_engine: 'Sharp v0.33.5'
        },
        environment: process.env.NODE_ENV || 'development'
    });
};

// Health Check Endpoints
app.get('/health', handleHealthCheck);
app.get('/healthz', handleHealthCheck);
app.get('/api/v1/health', handleHealthCheck);
app.get('/api/health', handleHealthCheck);

// Serve static frontend UI files from public/
app.use(express.static(path.join(__dirname, '../public')));

// Serve checkpoint_data.json sample file for Yamaha
app.get('/checkpoint_data.json', (req, res) => {
    const filePath = path.join(__dirname, '../checkpoint_data.json');
    res.sendFile(filePath, (err) => {
        if (err) {
            res.status(404).json({ error: 'Not Found', message: 'File checkpoint_data.json tidak ditemukan.' });
        }
    });
});

// Serve toyota_checkpoint_data.json sample file for Toyota
app.get('/toyota_checkpoint_data.json', (req, res) => {
    const filePath = path.join(__dirname, '../toyota_checkpoint_data.json');
    res.sendFile(filePath, (err) => {
        if (err) {
            res.status(404).json({ error: 'Not Found', message: 'File toyota_checkpoint_data.json tidak ditemukan.' });
        }
    });
});

// Serve asset_counting_data.json sample file for Asset Counting
app.get('/asset_counting_data.json', (req, res) => {
    const filePath = path.join(__dirname, '../asset_counting_data.json');
    res.sendFile(filePath, (err) => {
        if (err) {
            res.status(404).json({ error: 'Not Found', message: 'File asset_counting_data.json tidak ditemukan.' });
        }
    });
});

// Serve dummy_asset_counting_15.json
app.get('/dummy_asset_counting_15.json', (req, res) => {
    const filePath = path.join(__dirname, '../dummy_asset_counting_15.json');
    res.sendFile(filePath, (err) => {
        if (err) {
            res.status(404).json({ error: 'Not Found', message: 'File dummy_asset_counting_15.json tidak ditemukan.' });
        }
    });
});

// Serve dummy_checkpoint_15.json
app.get('/dummy_checkpoint_15.json', (req, res) => {
    const filePath = path.join(__dirname, '../dummy_checkpoint_15.json');
    res.sendFile(filePath, (err) => {
        if (err) {
            res.status(404).json({ error: 'Not Found', message: 'File dummy_checkpoint_15.json tidak ditemukan.' });
        }
    });
});

// Serve Dokumentasi_Spesifikasi_Mapping_Asset_Counting.xlsx
app.get('/Dokumentasi_Spesifikasi_Mapping_Asset_Counting.xlsx', (req, res) => {
    const filePath = path.join(__dirname, '../Dokumentasi_Spesifikasi_Mapping_Asset_Counting.xlsx');
    res.sendFile(filePath, (err) => {
        if (err) {
            res.status(404).json({ error: 'Not Found', message: 'File dokumentasi tidak ditemukan.' });
        }
    });
});

// Helper function untuk handle request Excel generation
async function handleExcelGeneration(req, res, reportType = 'yamaha') {
    try {
        let items = req.body;
        let areaFilter = req.query.area || req.query.area_filter || null;
        let sectionFilter = req.query.section || req.query.section_filter || null;
        let subdetailFilter = req.query.subdetail || req.query.sub_section || req.query.sectiondtl || null;

        if (items && !Array.isArray(items) && typeof items === 'object') {
            areaFilter = areaFilter || items.area_filter || items.area;
            sectionFilter = sectionFilter || items.section_filter || items.section;
            subdetailFilter = subdetailFilter || items.subdetail_filter || items.subdetail;

            // Jika tipe asset counting mengirim object root
            if (reportType.includes('asset')) {
                // Biarkan items utuh agar meta branch & fiscal_year terbaca di normalizer
            } else if (Array.isArray(items.data)) {
                items = items.data;
            } else {
                items = [items];
            }
        }

        const excelBuffer = await generateCheckpointExcel(items, reportType, areaFilter, sectionFilter, subdetailFilter);

        let filename = 'Laporan_Export.xlsx';
        const formattedReportType = String(reportType).toUpperCase().replace(/-/g, '_');

        if (reportType.includes('asset')) {
            const branchName = items.office_branch_name || items.office_name || items.branch_name || items.cabang || items.branch || 'AMBON';
            const year = items.fiscal_year || (items.doc_date ? new Date(items.doc_date).getFullYear() : (items.periode_start ? new Date(items.periode_start).getFullYear() : new Date().getFullYear()));
            filename = `Form_Asset_Counting_${String(branchName).toUpperCase()}_${year}.xlsx`;
        } else {
            const firstDoc = (Array.isArray(items) ? items[0]?.doc_num : items?.doc_num) || 'EXPORT';
            const safeDocNum = String(firstDoc).replace(/[/\\?%*:|"<>]/g, '_');
            
            let areaSuffix = '';
            if (areaFilter && areaFilter !== 'ALL') areaSuffix += `_${String(areaFilter).replace(/[/\\?%*:|"<>]/g, '_')}`;

            filename = `Laporan_Checkpoint_${formattedReportType}${areaSuffix}_${safeDocNum}.xlsx`;
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

        return res.send(excelBuffer);
    } catch (err) {
        console.error(`Error generating excel (${reportType}):`, err);
        return res.status(500).json({
            error: 'Internal Server Error',
            message: `Gagal memproses file Excel: ${err.message}`
        });
    }
}

// Endpoint spesifik YAMAHA: POST /api/v1/generate-excel/yamaha
app.post('/api/v1/generate-excel/yamaha', async (req, res) => {
    return handleExcelGeneration(req, res, 'yamaha');
});

// Endpoint spesifik TOYOTA: POST /api/v1/generate-excel/toyota
app.post('/api/v1/generate-excel/toyota', async (req, res) => {
    return handleExcelGeneration(req, res, 'toyota');
});

// Endpoint spesifik ASSET COUNTING: POST /api/v1/generate-excel/asset-counting
app.post('/api/v1/generate-excel/asset-counting', async (req, res) => {
    return handleExcelGeneration(req, res, 'asset-counting');
});

// Endpoint spesifik ASSET: POST /api/v1/generate-excel/asset
app.post('/api/v1/generate-excel/asset', async (req, res) => {
    return handleExcelGeneration(req, res, 'asset-counting');
});

// Endpoint dinamis untuk tipe report lain: POST /api/v1/generate-excel/:reportType
app.post('/api/v1/generate-excel/:reportType', async (req, res) => {
    const { reportType } = req.params;
    return handleExcelGeneration(req, res, reportType);
});

// Endpoint default fallback: POST /api/v1/generate-excel
app.post('/api/v1/generate-excel', async (req, res) => {
    return handleExcelGeneration(req, res, 'yamaha');
});

// 404 JSON Fallback Handler
app.use((req, res) => {
    res.status(404).json({
        error: 'Not Found',
        message: `Endpoint ${req.method} ${req.originalUrl} tidak ditemukan.`
    });
});

// 500 Global Error Handler
app.use((err, req, res, next) => {
    console.error('Unhandled Global Error:', err);
    res.status(500).json({
        error: 'Internal Server Error',
        message: err.message || 'Terjadi kesalahan internal pada server.'
    });
});

// Start server jika dijalankan langsung
if (require.main === module) {
    app.listen(PORT, () => {
        console.log(`🚀 Node.js Express API & Web UI running on http://localhost:${PORT}`);
        console.log(`📌 Yamaha Endpoint: POST http://localhost:${PORT}/api/v1/generate-excel/yamaha`);
        console.log(`📌 Toyota Endpoint: POST http://localhost:${PORT}/api/v1/generate-excel/toyota`);
        console.log(`📌 Asset Counting Endpoint: POST http://localhost:${PORT}/api/v1/generate-excel/asset-counting`);
        console.log(`💚 Health Check Endpoint: GET http://localhost:${PORT}/health`);
    });
}

module.exports = app;
