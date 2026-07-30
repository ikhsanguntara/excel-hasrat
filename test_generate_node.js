const fs = require('fs');
const path = require('path');
const { generateCheckpointExcel } = require('./src/excelGenerator');

async function testGeneration() {
    const jsonPath = path.join(__dirname, 'checkpoint_data.json');
    const outputPath = path.join(__dirname, 'laporan_checkpoint_node_sample.xlsx');

    if (!fs.existsSync(jsonPath)) {
        console.error(`Error: File ${jsonPath} tidak ditemukan!`);
        return;
    }

    const rawData = fs.readFileSync(jsonPath, 'utf-8');
    const data = JSON.parse(rawData);

    console.log(`Membaca ${data.length} item dari ${jsonPath}...`);

    const excelBuffer = await generateCheckpointExcel(data);

    fs.writeFileSync(outputPath, excelBuffer);

    console.log(`SUCCESS! File Excel Node.js berhasil dibuat di: ${outputPath}`);
    console.log(`Ukuran file: ${fs.statSync(outputPath).size} bytes`);
}

testGeneration().catch(err => {
    console.error("FAILED testing Node.js generation:", err);
});
