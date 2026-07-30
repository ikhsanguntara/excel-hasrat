document.addEventListener('DOMContentLoaded', () => {
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const browseBtn = document.getElementById('browseBtn');
    const filePreviewCard = document.getElementById('filePreviewCard');
    const fileNameEl = document.getElementById('fileName');
    const fileSizeEl = document.getElementById('fileSize');
    const removeFileBtn = document.getElementById('removeFileBtn');

    // Cascading Filter elements
    const areaSelect = document.getElementById('areaSelect');
    const sectionSelect = document.getElementById('sectionSelect');
    const subdetailSelect = document.getElementById('subdetailSelect');

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

        const first = items[0] || {};
        metricDocNum.textContent = first.doc_num || '-';
        metricCount.textContent = items.length;
        metricGroup.textContent = first.product_group || 'YAMAHA';

        // 1. Populate Area Options
        const areaSet = new Set(items.map(i => i.area_name).filter(Boolean));
        areaSelect.innerHTML = '<option value="ALL">Semua Area (All Areas)</option>';
        areaSet.forEach(areaName => {
            const opt = document.createElement('option');
            opt.value = areaName;
            opt.textContent = areaName;
            areaSelect.appendChild(opt);
        });

        metricAreas.textContent = areaSet.size > 0 ? areaSet.size : 1;

        // Populate Sections and Subdetails initially
        updateSectionOptions();

        filePreviewCard.classList.remove('hidden');
        generateBtn.disabled = false;
    }

    // Cascading Logic: Update Section Options based on Area Selection
    function updateSectionOptions() {
        if (!currentJsonData) return;
        const selectedArea = areaSelect.value;
        
        let filtered = currentJsonData;
        if (selectedArea && selectedArea !== 'ALL') {
            filtered = filtered.filter(i => i.area_name === selectedArea);
        }

        const sectionSet = new Set(filtered.map(i => i.section_name).filter(Boolean));
        sectionSelect.innerHTML = '<option value="ALL">Semua Section (All Sections)</option>';
        sectionSet.forEach(secName => {
            const opt = document.createElement('option');
            opt.value = secName;
            opt.textContent = secName;
            sectionSelect.appendChild(opt);
        });

        updateSubdetailOptions();
    }

    // Cascading Logic: Update Subdetail Options based on Section Selection
    function updateSubdetailOptions() {
        if (!currentJsonData) return;
        const selectedArea = areaSelect.value;
        const selectedSec = sectionSelect.value;

        let filtered = currentJsonData;
        if (selectedArea && selectedArea !== 'ALL') {
            filtered = filtered.filter(i => i.area_name === selectedArea);
        }
        if (selectedSec && selectedSec !== 'ALL') {
            filtered = filtered.filter(i => i.section_name === selectedSec);
        }

        const subSet = new Set(filtered.map(i => i.sectiondtl_name).filter(Boolean));
        subdetailSelect.innerHTML = '<option value="ALL">Semua Sub-bagian (All)</option>';
        subSet.forEach(subName => {
            const opt = document.createElement('option');
            opt.value = subName;
            opt.textContent = subName;
            subdetailSelect.appendChild(opt);
        });
    }

    areaSelect.addEventListener('change', updateSectionOptions);
    sectionSelect.addEventListener('change', updateSubdetailOptions);

    removeFileBtn.addEventListener('click', () => {
        currentJsonData = null;
        currentFileName = "";
        fileInput.value = "";
        areaSelect.innerHTML = '<option value="ALL">Semua Area (All Areas)</option>';
        sectionSelect.innerHTML = '<option value="ALL">Semua Section (All Sections)</option>';
        subdetailSelect.innerHTML = '<option value="ALL">Semua Sub-bagian (All)</option>';
        filePreviewCard.classList.add('hidden');
        generateBtn.disabled = true;
        hideToast();
    });

    // 3. Gunakan Contoh JSON
    useSampleBtn.addEventListener('click', async () => {
        try {
            const resp = await fetch('/checkpoint_data.json');
            if (!resp.ok) {
                throw new Error(`Gagal mengambil file contoh (HTTP ${resp.status})`);
            }
            const contentType = resp.headers.get('content-type');
            if (contentType && !contentType.includes('json')) {
                throw new Error('Server mengembalikan file non-JSON (HTML 404/Page)');
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

    // 4. Generate & Download Excel (dengan Cascading Filters)
    generateBtn.addEventListener('click', async () => {
        if (!currentJsonData) return;

        setLoading(true);
        hideToast();

        const selectedArea = areaSelect.value;
        const selectedSec = sectionSelect.value;
        const selectedSub = subdetailSelect.value;

        const params = new URLSearchParams();
        if (selectedArea && selectedArea !== 'ALL') params.append('area', selectedArea);
        if (selectedSec && selectedSec !== 'ALL') params.append('section', selectedSec);
        if (selectedSub && selectedSub !== 'ALL') params.append('subdetail', selectedSub);

        let apiUrl = '/api/v1/generate-excel/yamaha';
        if (params.toString()) {
            apiUrl += `?${params.toString()}`;
        }

        try {
            const response = await fetch(apiUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(currentJsonData)
            });

            if (!response.ok) {
                let errMsg = `HTTP Status ${response.status}: ${response.statusText}`;
                try {
                    const contentType = response.headers.get('content-type');
                    if (contentType && contentType.includes('application/json')) {
                        const errData = await response.json();
                        errMsg = errData.message || errData.error || errMsg;
                    } else {
                        const text = await response.text();
                        if (text && !text.includes('<!DOCTYPE')) {
                            errMsg = text.substring(0, 150);
                        }
                    }
                } catch (e) {
                    // Ignore parse error fallback
                }
                throw new Error(errMsg);
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
