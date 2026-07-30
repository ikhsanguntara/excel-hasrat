const { getGenerator, registerGenerator } = require('./generators');

/**
 * Main entry point generator Excel.
 * Menerima list data checkpoint, tipe laporan, dan opsional areaFilter.
 */
async function generateCheckpointExcel(items, reportType = 'yamaha', areaFilter = null) {
    const generator = getGenerator(reportType);
    return await generator(items, areaFilter);
}

module.exports = {
    generateCheckpointExcel,
    getGenerator,
    registerGenerator
};
