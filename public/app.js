document.addEventListener('DOMContentLoaded', () => {
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const browseBtn = document.getElementById('browseBtn');
    const filePreviewCard = document.getElementById('filePreviewCard');
    const fileNameEl = document.getElementById('fileName');
    const fileSizeEl = document.getElementById('fileSize');
    const removeFileBtn = document.getElementById('removeFileBtn');
    const generateBtn = document.getElementById('generateBtn');
    const useSampleBtn = document.getElementById('useSampleBtn');
    const spinner = document.getElementById('spinner');
    const toast = document.getElementById('toast');
    const toastMessage = document.getElementById('toastMessage');

    // Metrics elements
    const metricDocNum = document.getElementById('metricDocNum');
    const metricCount = document.getElementById('metricCount');
    const metricAreas = document.getElementById('metricAreas');
    const metricGroup = document.getElementById('metricGroup');

    let currentJsonData = null;
    let currentFileName = "";

    // 1. Drag and Drop events
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => dropzone.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => dropzone.classList.remove('dragover'), false);
    });

    dropzone.addEventListener('drop', handleDrop, false);

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleFileSelect(files[0]);
        }
    }

    browseBtn.addEventListener('click', () => fileInput.click());
    dropzone.addEventListener('click', (e) => {
        if (e.target !== browseBtn) fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });

    // 2. File Processing
    function handleFileSelect(file) {
        if (!file.name.endsWith('.json')) {
            showToast('Harap pilih file dengan ekstensi .json', 'error');
            return;
        }

        currentFileName = file.name;
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const jsonText = e.target.result;
                const parsed = JSON.parse(jsonText);
                
                let dataArray = [];
                if (Array.isArray(parsed)) {
                    dataArray = parsed;
                } else if (parsed && typeof parsed === 'object' && Array.isArray(parsed.data)) {
                    dataArray = parsed.data;
                } else if (parsed && typeof parsed === 'object') {
                    dataArray = [parsed];
                }

                if (dataArray.length === 0) {
                    showToast('File JSON tidak berisi data array checkpoint yang valid.', 'error');
                    return;
                }

                currentJsonData = dataArray;
                updateUIWithFile(file.name, file.size, dataArray);
                showToast('File JSON berhasil dimuat!', 'success');
            } catch (err) {
                showToast('Format JSON tidak valid. Periksa sintaks file.', 'error');
                console.error(err);
            }
        };

        reader.readAsText(file);
    }

    function updateUIWithFile(filename, bytes, items) {
        fileNameEl.textContent = filename;
        fileSizeEl.textContent = `${formatBytes(bytes)} • ${items.length} Checkpoint Items`;

        // Compute metrics
        const first = items[0] || {};
        metricDocNum.textContent = first.doc_num || '-';
        metricCount.textContent = items.length;
        
        const areas = new Set(items.map(i => i.area_name).filter(Boolean));
        metricAreas.textContent = areas.size > 0 ? areas.size : 1;
        metricGroup.textContent = first.product_group || 'YAMAHA';

        filePreviewCard.classList.remove('hidden');
        generateBtn.disabled = false;
    }

    removeFileBtn.addEventListener('click', () => {
        currentJsonData = null;
        currentFileName = "";
        fileInput.value = "";
        filePreviewCard.classList.add('hidden');
        generateBtn.disabled = true;
        hideToast();
    });

    // 3. Gunakan Contoh JSON
    useSampleBtn.addEventListener('click', async () => {
        try {
            const resp = await fetch('/checkpoint_data.json');
            if (!resp.ok) {
                throw new Error('Gagal mengambil file sampel checkpoint_data.json');
            }
            const data = await resp.json();
            currentJsonData = data;
            currentFileName = 'checkpoint_data.json';

            updateUIWithFile('checkpoint_data.json', JSON.stringify(data).length, data);
            showToast('Sampel checkpoint_data.json berhasil dimuat!', 'success');
        } catch (err) {
            showToast('Gagal memuat file contoh: ' + err.message, 'error');
        }
    });

    // 4. Generate & Download Excel (Menggunakan /api/v1/generate-excel/yamaha)
    generateBtn.addEventListener('click', async () => {
        if (!currentJsonData) return;

        setLoading(true);
        hideToast();

        try {
            const response = await fetch('/api/v1/generate-excel/yamaha', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(currentJsonData)
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.message || 'Gagal generate file Excel.');
            }

            // Ambil filename dari Header Content-Disposition jika ada
            let filename = 'Laporan_Checkpoint_YAMAHA.xlsx';
            const disposition = response.headers.get('Content-Disposition');
            if (disposition && disposition.includes('filename=')) {
                const match = disposition.match(/filename="?([^"]+)"?/);
                if (match && match[1]) {
                    filename = match[1];
                }
            }

            const blob = await response.blob();
            const downloadUrl = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(downloadUrl);

            showToast('File Excel Yamaha berhasil dibuat & didownload!', 'success');
        } catch (err) {
            showToast('Error: ' + err.message, 'error');
            console.error(err);
        } finally {
            setLoading(false);
        }
    });

    function setLoading(isLoading) {
        if (isLoading) {
            generateBtn.disabled = true;
            spinner.classList.remove('hidden');
        } else {
            generateBtn.disabled = false;
            spinner.classList.add('hidden');
        }
    }

    function showToast(msg, type = 'success') {
        toastMessage.textContent = msg;
        toast.className = `toast ${type}`;
    }

    function hideToast() {
        toast.className = 'toast hidden';
    }

    function formatBytes(bytes, decimals = 1) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }
});
