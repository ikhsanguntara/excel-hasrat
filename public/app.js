document.addEventListener('DOMContentLoaded', () => {
    const btnTypeYamaha = document.getElementById('btnTypeYamaha');
    const btnTypeToyota = document.getElementById('btnTypeToyota');

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
    const sampleBtnText = document.getElementById('sampleBtnText');
    const spinner = document.getElementById('spinner');
    const toast = document.getElementById('toast');
    const toastMessage = document.getElementById('toastMessage');

    // Metrics elements
    const metricDocNum = document.getElementById('metricDocNum');
    const metricCount = document.getElementById('metricCount');
    const metricAreas = document.getElementById('metricAreas');
    const metricGroup = document.getElementById('metricGroup');

    let currentReportType = 'yamaha'; // 'yamaha' or 'toyota'
    let currentJsonData = null;
    let currentFileName = "";

    // 0. Toggle Tipe Laporan (Yamaha vs Toyota)
    btnTypeYamaha.addEventListener('click', () => setReportType('yamaha'));
    btnTypeToyota.addEventListener('click', () => setReportType('toyota'));

    function setReportType(type) {
        currentReportType = type;
        if (type === 'yamaha') {
            btnTypeYamaha.classList.add('active');
            btnTypeToyota.classList.remove('active');
            sampleBtnText.textContent = 'Gunakan Contoh JSON Yamaha';
        } else {
            btnTypeToyota.classList.add('active');
            btnTypeYamaha.classList.remove('active');
            sampleBtnText.textContent = 'Gunakan Contoh JSON Toyota';
        }

        // Reset data saat ganti tipe jika belum ada file khusus
        if (currentJsonData) {
            updateUIWithFile(currentFileName, JSON.stringify(currentJsonData).length, currentJsonData);
        }
    }

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

    // Helper untuk me-flatten data jika data berupa Toyota nested
    function extractFlatItems(parsed) {
        let items = [];

        if (Array.isArray(parsed)) {
            parsed.forEach(entry => {
                if (entry.sections && Array.isArray(entry.sections)) {
                    // Toyota nested structure
                    const areaName = entry.name || 'Showroom Toyota';
                    entry.sections.forEach(sec => {
                        const secName = sec.name || 'Umum';
                        (sec.items || []).forEach(it => {
                            items.push({
                                area_name: areaName,
                                section_name: secName,
                                sectiondtl_name: '',
                                checkpoint_name: it.checkPoint || '',
                                result: it.hasilPenilaian || '',
                                solution: it.solution || '',
                                img_path: it.hasilFoto || ''
                            });
                        });
                    });
                } else {
                    items.push(entry);
                }
            });
        } else if (parsed && typeof parsed === 'object' && Array.isArray(parsed.data)) {
            items = parsed.data;
        } else if (parsed && typeof parsed === 'object') {
            items = [parsed];
        }

        return items;
    }

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

                // Auto-detect tipe jika nama file mengandung toyota
                if (file.name.toLowerCase().includes('toyota')) {
                    setReportType('toyota');
                } else if (file.name.toLowerCase().includes('yamaha')) {
                    setReportType('yamaha');
                }

                currentJsonData = parsed;
                updateUIWithFile(file.name, file.size, parsed);
                showToast(`File JSON (${currentReportType.toUpperCase()}) berhasil dimuat!`, 'success');
            } catch (err) {
                showToast('Format JSON tidak valid. Periksa sintaks file.', 'error');
                console.error(err);
            }
        };

        reader.readAsText(file);
    }

    function updateUIWithFile(filename, bytes, rawData) {
        const flatItems = extractFlatItems(rawData);

        fileNameEl.textContent = filename;
        fileSizeEl.textContent = `${formatBytes(bytes)} • ${flatItems.length} Checkpoint Items`;

        const first = flatItems[0] || {};
        metricDocNum.textContent = first.doc_num || (currentReportType === 'toyota' ? 'TYT/2026/001' : 'YMH/2026/001');
        metricCount.textContent = flatItems.length;
        metricGroup.textContent = currentReportType.toUpperCase();

        // 1. Populate Area Options
        const areaSet = new Set(flatItems.map(i => i.area_name || i.name).filter(Boolean));
        areaSelect.innerHTML = '<option value="ALL">Semua Area (All Areas)</option>';
        areaSet.forEach(areaName => {
            const opt = document.createElement('option');
            opt.value = areaName;
            opt.textContent = areaName;
            areaSelect.appendChild(opt);
        });

        metricAreas.textContent = areaSet.size > 0 ? areaSet.size : 1;

        // Populate Sections and Subdetails initially
        updateSectionOptions(flatItems);

        filePreviewCard.classList.remove('hidden');
        generateBtn.disabled = false;
    }

    // Cascading Logic: Update Section Options based on Area Selection
    function updateSectionOptions(flatItemsPassed = null) {
        const flatItems = flatItemsPassed || extractFlatItems(currentJsonData);
        if (!flatItems || flatItems.length === 0) return;

        const selectedArea = areaSelect.value;
        
        let filtered = flatItems;
        if (selectedArea && selectedArea !== 'ALL') {
            filtered = filtered.filter(i => (i.area_name || i.name) === selectedArea);
        }

        const sectionSet = new Set(filtered.map(i => i.section_name).filter(Boolean));
        sectionSelect.innerHTML = '<option value="ALL">Semua Section (All Sections)</option>';
        sectionSet.forEach(secName => {
            const opt = document.createElement('option');
            opt.value = secName;
            opt.textContent = secName;
            sectionSelect.appendChild(opt);
        });

        updateSubdetailOptions(flatItems);
    }

    // Cascading Logic: Update Subdetail Options based on Section Selection
    function updateSubdetailOptions(flatItemsPassed = null) {
        const flatItems = flatItemsPassed || extractFlatItems(currentJsonData);
        if (!flatItems || flatItems.length === 0) return;

        const selectedArea = areaSelect.value;
        const selectedSec = sectionSelect.value;

        let filtered = flatItems;
        if (selectedArea && selectedArea !== 'ALL') {
            filtered = filtered.filter(i => (i.area_name || i.name) === selectedArea);
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

    areaSelect.addEventListener('change', () => updateSectionOptions());
    sectionSelect.addEventListener('change', () => updateSubdetailOptions());

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
        const sampleUrl = currentReportType === 'toyota' ? '/toyota_checkpoint_data.json' : '/checkpoint_data.json';
        const sampleName = currentReportType === 'toyota' ? 'toyota_checkpoint_data.json' : 'checkpoint_data.json';

        try {
            const resp = await fetch(sampleUrl);
            if (!resp.ok) {
                throw new Error(`Gagal mengambil file contoh (HTTP ${resp.status})`);
            }
            const contentType = resp.headers.get('content-type');
            if (contentType && !contentType.includes('json')) {
                throw new Error('Server mengembalikan file non-JSON (HTML 404/Page)');
            }
            const data = await resp.json();
            currentJsonData = data;
            currentFileName = sampleName;

            updateUIWithFile(sampleName, JSON.stringify(data).length, data);
            showToast(`Sampel ${sampleName} berhasil dimuat!`, 'success');
        } catch (err) {
            showToast('Gagal memuat file contoh: ' + err.message, 'error');
        }
    });

    // 4. Generate & Download Excel (Yamaha / Toyota API Endpoint)
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

        let apiUrl = `/api/v1/generate-excel/${currentReportType}`;
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
            let filename = `Laporan_Checkpoint_${currentReportType.toUpperCase()}.xlsx`;
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

            showToast(`File Excel ${currentReportType.toUpperCase()} berhasil dibuat & didownload!`, 'success');
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
