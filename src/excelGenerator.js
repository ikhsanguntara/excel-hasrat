const { getGenerator, registerGenerator } = require('./generators');

/**
 * Main entry point generator Excel.
 * Menerima list data checkpoint dan tipe laporan (default: 'yamaha').
 */
async function generateCheckpointExcel(items, reportType = 'yamaha') {
    const generator = getGenerator(reportType);
    return await generator(items);
}

module.exports = {
    generateCheckpointExcel,
    getGenerator,
    registerGenerator
};
