const path = require('path');
const express = require('express');
const { generateCheckpointExcel } = require('./excelGenerator');

const app = express();
const PORT = process.env.PORT || 8000;

// Middleware for JSON body parsing (limit 50mb for large JSONs)
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

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
            const branchName = items.branch_name || items.cabang || items.branch || 'AMBON';
            filename = `Form_Asset_Counting_${String(branchName).toUpperCase()}_${new Date().getFullYear()}.xlsx`;
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
    });
}

module.exports = app;
