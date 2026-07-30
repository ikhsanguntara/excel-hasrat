const { generateYamahaExcel } = require('./yamahaGenerator');

/**
 * Registry untuk menyimpan berbagai jenis generator laporan Excel.
 * Jika di masa depan ada format baru (misal: 'honda', 'toyota'), 
 * developer cukup menambahkan file generator baru di folder `src/generators/` dan mengdaftarkannya di sini.
 */
const generatorRegistry = {
    'yamaha': generateYamahaExcel,
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
