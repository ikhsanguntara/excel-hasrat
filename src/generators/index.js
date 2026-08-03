const { generateYamahaExcel } = require('./yamahaGenerator');
const { generateToyotaExcel } = require('./toyotaGenerator');

/**
 * Registry untuk menyimpan berbagai jenis generator laporan Excel.
 * Mendukung format: 'yamaha', 'toyota', dan 'default'.
 */
const generatorRegistry = {
    'yamaha': generateYamahaExcel,
    'toyota': generateToyotaExcel,
    'default': generateYamahaExcel
};

/**
 * Mengambil fungsi generator berdasarkan nama reportType.
 */
function getGenerator(reportType) {
    const key = String(reportType || 'yamaha').toLowerCase().trim();
    if (generatorRegistry[key]) {
        return generatorRegistry[key];
    }
    // Return default generator jika tipe tidak ditemukan
    return generatorRegistry['default'];
}

/**
 * Mendaftarkan generator baru secara dinamis
 */
function registerGenerator(reportType, generatorFunction) {
    const key = String(reportType).toLowerCase().trim();
    generatorRegistry[key] = generatorFunction;
}

module.exports = {
    getGenerator,
    registerGenerator,
    generatorRegistry
};
