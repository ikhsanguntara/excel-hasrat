const { getGenerator, registerGenerator } = require('./generators');

/**
 * Main entry point generator Excel.
 * Menerima list data checkpoint, tipe laporan, dan opsional filter bertingkat (area, section, subdetail).
 */
async function generateCheckpointExcel(items, reportType = 'yamaha', areaFilter = null, sectionFilter = null, subdetailFilter = null) {
    const generator = getGenerator(reportType);
    return await generator(items, areaFilter, sectionFilter, subdetailFilter);
}

module.exports = {
    generateCheckpointExcel,
    getGenerator,
    registerGenerator
};
