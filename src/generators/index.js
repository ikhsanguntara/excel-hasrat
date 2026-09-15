const { generateYamahaExcel } = require('./yamahaGenerator');
const { generateToyotaExcel } = require('./toyotaGenerator');
const { generateAssetCountingExcel } = require('./assetCountingGenerator');

/**
 * Registry untuk menyimpan berbagai jenis generator laporan Excel.
 * Mendukung format: 'yamaha', 'toyota', 'asset-counting', 'asset', dan 'default'.
 */
const generatorRegistry = {
    'yamaha': generateYamahaExcel,
    'toyota': generateToyotaExcel,
    'asset-counting': generateAssetCountingExcel,
    'asset_counting': generateAssetCountingExcel,
    'asset': generateAssetCountingExcel,
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
