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

// Serve checkpoint_data.json sample file for the UI
app.get('/checkpoint_data.json', (req, res) => {
    res.sendFile(path.join(__dirname, '../checkpoint_data.json'));
});

// Helper function untuk handle request Excel generation
async function handleExcelGeneration(req, res, reportType = 'yamaha') {
    try {
        let items = req.body;

        if (items && !Array.isArray(items) && typeof items === 'object') {
            if (Array.isArray(items.data)) {
                items = items.data;
            } else {
                items = [items];
            }
        }

        if (!Array.isArray(items) || items.length === 0) {
            return res.status(400).json({
                error: 'Bad Request',
                message: 'Data JSON tidak boleh kosong dan harus berupa Array atau Object dengan key "data".'
            });
        }

        const excelBuffer = await generateCheckpointExcel(items, reportType);

        const firstDoc = items[0].doc_num || 'EXPORT';
        const safeDocNum = String(firstDoc).replace(/[/\\?%*:|"<>]/g, '_');
        const formattedReportType = String(reportType).toUpperCase();
        const filename = `Laporan_Checkpoint_${formattedReportType}_${safeDocNum}.xlsx`;

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

// Endpoint spesifik untuk YAMAHA: POST /api/v1/generate-excel/yamaha
app.post('/api/v1/generate-excel/yamaha', async (req, res) => {
    return handleExcelGeneration(req, res, 'yamaha');
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

// Start server jika dijalankan langsung
if (require.main === module) {
    app.listen(PORT, () => {
        console.log(`🚀 Node.js Express API & Web UI running on http://localhost:${PORT}`);
        console.log(`📌 Yamaha Endpoint: POST http://localhost:${PORT}/api/v1/generate-excel/yamaha`);
    });
}

module.exports = app;
