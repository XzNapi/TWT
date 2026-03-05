// ============================================
// CERTIFICATE GENERATOR - MAIN SCRIPT
// ============================================

// Inisialisasi PDF.js Worker
if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
}

// ============================================
// DOM ELEMENTS - GROUPED BY FUNCTION
// ============================================

// Upload & Sections
const uploadInput = document.getElementById('design-upload');
const editorSection = document.getElementById('editor-section');
const bulkSection = document.getElementById('bulk-section');
const uploadSection = document.getElementById('upload-section');

// Canvas
const canvas = document.getElementById('sertifikat-canvas');
const ctx = canvas.getContext('2d', { willReadFrequently: true });

// Split Pane Elements
const leftPane = document.getElementById('editor-left-pane');
const resizer = document.getElementById('split-resizer');
const splitContainer = document.querySelector('.editor-split-container');

// Text Editing Controls
const textEditControls = document.getElementById('text-edit-controls');
const addTextBtn = document.getElementById('add-text-btn');
const deleteTextBtn = document.getElementById('delete-text-btn');
const deleteTextContainer = document.getElementById('delete-text-container');

// Form Inputs
const dataLinkSelect = document.getElementById('data-link-select');
const textInput = document.getElementById('text-input');
const fontFamilySelect = document.getElementById('font-family');
const fontSizeInput = document.getElementById('font-size');
const fontColorInput = document.getElementById('font-color');
const fontColorHexInput = document.getElementById('font-color-hex');
const colorPickerBtn = document.getElementById('color-picker-btn');
const fontBoldBtn = document.getElementById('font-bold');
const fontItalicBtn = document.getElementById('font-italic');
const fontAlignSelect = document.getElementById('font-align');
const textTransformSelect = document.getElementById('text-transform');

// Font Management
const addFontBtn = document.getElementById('add-font-btn');
const newFontNameInput = document.getElementById('new-font-name');
const newFontFileInput = document.getElementById('new-font-file');

// Position Controls (Nudge)
const moveUpBtn = document.getElementById('move-up');
const moveDownBtn = document.getElementById('move-down');
const moveLeftBtn = document.getElementById('move-left');
const moveRightBtn = document.getElementById('move-right');
const NUDGE_AMOUNT = 5;

// Template Management
const saveTemplateBtn = document.getElementById('save-template-btn');
const loadTemplateBtn = document.getElementById('load-template-btn');
const loadTemplateInput = document.getElementById('load-template-input');
const downloadBtn = document.getElementById('download-btn');

// Bulk Generation
const bulkFileInput = document.getElementById('bulk-file-upload');
const gdocLinkInput = document.getElementById('gdoc-link');
const fetchGdocBtn = document.getElementById('fetch-gdoc-btn');
const bulkFormatSelect = document.getElementById('bulk-format');
const generateBulkBtn = document.getElementById('generate-bulk-btn');
const bulkStatusDiv = document.getElementById('bulk-status');
const zipCheckbox = document.getElementById('download-as-zip');

// ============================================
// STATE MANAGEMENT
// ============================================

const state = {
    backgroundImage: null,
    textFields: [],
    selectedTextId: null,
    nextTextId: 0,
    dataList: [],
    dataHeaders: [],
    isDragging: false,
    dragStartOffset: { x: 0, y: 0 },
    isResizing: false
};

// ============================================
// SPLIT PANE FUNCTIONALITY
// ============================================

function initSplitPane() {
    if (!resizer || !leftPane || !splitContainer) return;

    const startResize = (e) => {
        state.isResizing = true;
        resizer.classList.add('resizing');
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault();
    };

    const doResize = (e) => {
        if (!state.isResizing) return;
        
        const clientX = e.touches ? e.touches[0].clientX : e.clientX;
        const containerRect = splitContainer.getBoundingClientRect();
        const newWidth = clientX - containerRect.left;
        
        // Constraints: min 280px, max 500px
        const clampedWidth = Math.max(280, Math.min(500, newWidth));
        leftPane.style.width = `${clampedWidth}px`;
    };

    const stopResize = () => {
        if (!state.isResizing) return;
        state.isResizing = false;
        resizer.classList.remove('resizing');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
    };

    // Mouse events
    resizer.addEventListener('mousedown', startResize);
    document.addEventListener('mousemove', doResize);
    document.addEventListener('mouseup', stopResize);

    // Touch events for mobile
    resizer.addEventListener('touchstart', startResize, { passive: false });
    document.addEventListener('touchmove', doResize, { passive: false });
    document.addEventListener('touchend', stopResize);
}

// ============================================
// CANVAS & RENDERING
// ============================================

function redrawCanvas() {
    if (!state.backgroundImage?.src) return;

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(state.backgroundImage, 0, 0, canvas.width, canvas.height);

    state.textFields.forEach(field => {
        renderTextField(field);
    });
}

function renderTextField(field) {
    // Build font style
    let style = "";
    if (field.isItalic) style += "italic ";
    if (field.isBold) style += "bold ";
    
    ctx.font = `${style}${field.size}px "${field.family}"`;
    ctx.fillStyle = field.color;
    ctx.textAlign = field.align;
    ctx.textBaseline = 'middle';

    const textToDraw = field.dataLink ? `[${field.dataLink}]` : field.text;
    updateTextFieldBoundingBox(field, textToDraw);
    ctx.fillText(textToDraw, field.x, field.y);

    // Draw selection box if selected
    if (field.id === state.selectedTextId) {
        drawSelectionBox(field);
    }
}

function drawSelectionBox(field) {
    ctx.strokeStyle = 'rgba(37, 99, 235, 0.8)';
    ctx.lineWidth = 2;
    ctx.setLineDash([6, 4]);
    ctx.strokeRect(
        field.boundingBox.x, 
        field.boundingBox.y, 
        field.boundingBox.width, 
        field.boundingBox.height
    );
    ctx.setLineDash([]);
    ctx.lineWidth = 1;
}

function updateTextFieldBoundingBox(field, textToDraw) {
    ctx.font = `${field.isItalic ? 'italic ' : ''}${field.isBold ? 'bold ' : ''}${field.size}px "${field.family}"`;
    ctx.textAlign = field.align;
    ctx.textBaseline = 'middle';

    const metrics = ctx.measureText(textToDraw);
    const height = (metrics.actualBoundingBoxAscent || field.size * 0.8) + 
                   (metrics.actualBoundingBoxDescent || field.size * 0.2);
    const width = metrics.width;
    const paddingX = 10;
    const paddingY = 10;

    field.boundingBox = {
        width: width + (paddingX * 2),
        height: height + (paddingY * 2),
        x: field.align === 'center' ? field.x - (width / 2) - paddingX :
           field.align === 'right' ? field.x - width - paddingX :
           field.x - paddingX,
        y: field.y - (height / 2) - paddingY
    };
}

// ============================================
// TEXT FIELD MANAGEMENT
// ============================================

function addNewTextField(linkToData = null) {
    const defaultText = linkToData ? `[${linkToData}]` : "Teks Baru";
    const newField = {
        id: state.nextTextId++,
        text: defaultText,
        dataLink: linkToData,
        x: canvas.width / 2,
        y: (canvas.height / 2) + (state.textFields.length * 50),
        size: 50,
        family: 'Times New Roman',
        color: '#000000',
        isBold: false,
        isItalic: false,
        align: 'center',
        transform: 'none',
        boundingBox: {}
    };
    
    state.textFields.push(newField);
    selectTextField(newField.id);
}

function selectTextField(id) {
    state.selectedTextId = id;
    
    if (id === null) {
        textEditControls.classList.add('hidden');
        deleteTextContainer.classList.add('invisible');
    } else {
        textEditControls.classList.remove('hidden');
        deleteTextContainer.classList.remove('invisible');
        updateToolbarForSelected();
    }
    
    redrawCanvas();
}

function updateSelectedTextField(property, value) {
    const field = state.textFields.find(t => t.id === state.selectedTextId);
    if (!field) return;

    field[property] = value;

    if (property === 'dataLink') {
        if (value === "STATIC_TEXT") {
            field.dataLink = null;
            field.text = "Teks Statis";
        } else {
            field.dataLink = value;
            field.text = `[${value}]`;
        }
        updateToolbarForSelected();
    }
    
    if (property === 'text') {
        field.dataLink = null;
    }
    
    redrawCanvas();
}

function updateToolbarForSelected() {
    const field = state.textFields.find(t => t.id === state.selectedTextId);
    if (!field) return;

    // Update all inputs
    fontFamilySelect.value = field.family;
    fontSizeInput.value = field.size;
    fontColorInput.value = field.color;
    fontColorHexInput.value = field.color;
    fontBoldBtn.classList.toggle('active', field.isBold);
    fontItalicBtn.classList.toggle('active', field.isItalic);
    fontAlignSelect.value = field.align;
    textTransformSelect.value = field.transform || 'none';

    // Handle data link vs static text
    if (field.dataLink) {
        dataLinkSelect.value = field.dataLink;
        textInput.value = `[${field.dataLink}]`;
        textInput.disabled = true;
        textInput.classList.add('disabled');
    } else {
        dataLinkSelect.value = "STATIC_TEXT";
        textInput.value = field.text;
        textInput.disabled = false;
        textInput.classList.remove('disabled');
    }
}

function updateDataLinkDropdown() {
    dataLinkSelect.innerHTML = '';
    
    const staticOption = document.createElement('option');
    staticOption.value = "STATIC_TEXT";
    staticOption.textContent = "Teks Statis (Tidak Di-link)";
    dataLinkSelect.appendChild(staticOption);
    
    state.dataHeaders.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.textContent = header;
        dataLinkSelect.appendChild(option);
    });
}

// ============================================
// DATA PROCESSING
// ============================================

function toTitleCase(str) {
    if (!str) return "";
    return String(str).toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
}

function processNamaList(list, isGSheet = false, explicitHeaders = null) {
    state.dataList = [];
    state.dataHeaders = [];

    // Parse data based on source type
    if (isGSheet) {
        parseGSheetData(list);
    } else {
        state.dataList = list;
        state.dataHeaders = explicitHeaders || (list.length > 0 ? Object.keys(list[0]) : []);
    }
    
    if (state.dataList.length === 0) {
        showStatus("Tidak ada baris data yang ditemukan.", 'error');
        return;
    }

    // Check for duplicates
    const warnings = checkDuplicates();
    updateUIAfterDataLoad(warnings);
}

function parseGSheetData(csvContent) {
    const lines = csvContent.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    
    if (lines.length < 2) {
        throw new Error("File CSV kosong atau tidak valid");
    }
    
    const headers = lines[0].split(',').map(h => h.trim());
    state.dataHeaders = headers;
    
    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(',');
        const row = {};
        headers.forEach((header, index) => {
            row[header] = values[index] ? values[index].trim() : '';
        });
        state.dataList.push(row);
    }
}

function checkDuplicates() {
    const warnings = [];
    const headerToIndex = new Map(state.dataHeaders.map((h, i) => [h, String.fromCharCode(65 + i)]));

    state.dataHeaders.forEach(header => {
        const valueCounts = new Map();
        
        state.dataList.forEach(row => {
            const value = row[header];
            if (!value) return;
            
            const standardized = String(value).toLowerCase().trim();
            if (!standardized) return;
            
            const current = valueCounts.get(standardized);
            if (current) {
                current.count++;
            } else {
                valueCounts.set(standardized, { count: 1, originalValue: value });
            }
        });

        const duplicates = [];
        valueCounts.forEach((data, stdVal) => {
            if (data.count > 1) {
                duplicates.push(data.originalValue);
            }
        });
        
        if (duplicates.length > 0) {
            const cellLetter = headerToIndex.get(header) || header;
            warnings.push({
                cell: cellLetter,
                values: duplicates,
                count: duplicates.length
            });
        }
    });
    
    return warnings;
}

function updateUIAfterDataLoad(warnings) {
    // Build status message
    let message = '';
    let type = 'info';
    
    if (state.backgroundImage?.src) {
        message = `Siap menghasilkan ${state.dataList.length} sertifikat.`;
        type = 'success';
        generateBulkBtn.disabled = false;
    } else {
        message = `Data terbaca (${state.dataList.length} baris, ${state.dataHeaders.length} kolom). Silakan upload desain.`;
        type = 'info';
    }
    
    // Add warnings
    if (warnings.length > 0) {
        const warningHtml = warnings.map(w => 
            `<br><strong style="color: #e67e22;">⚠️ Peringatan:</strong> ` +
            `${w.count} duplikasi pada kolom <strong>${w.cell}</strong>: ${w.values.join(', ')}`
        ).join('');
        message += warningHtml;
    }
    
    showStatus(message, type);
    updateDataLinkDropdown();
    
    // Auto-add first text field if none exists
    if (state.textFields.length === 0) {
        addNewTextField(state.dataHeaders.length > 0 ? state.dataHeaders[0] : null);
    }
}

function showStatus(message, type = 'info') {
    bulkStatusDiv.innerHTML = message;
    bulkStatusDiv.className = 'status-message show ' + type;
    
    const colors = {
        success: '#059669',
        error: '#dc2626',
        info: '#2563eb'
    };
    bulkStatusDiv.style.color = colors[type] || colors.info;
}

// ============================================
// FILE UPLOAD HANDLERS
// ============================================

function handleDesignUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (file.type.startsWith('image/')) {
        loadImageFile(file);
    } else if (file.type === 'application/pdf') {
        loadPdfFile(file);
    } else {
        alert("Format file tidak didukung. Harap unggah .png, .jpeg, atau .pdf");
        uploadInput.value = '';
    }
}

function loadImageFile(file) {
    const reader = new FileReader();
    reader.onload = (event) => {
        const img = new Image();
        img.onload = () => {
            setBackgroundImage(img);
        };
        img.src = event.target.result;
    };
    reader.readAsDataURL(file);
}

function loadPdfFile(file) {
    if (typeof pdfjsLib === 'undefined') {
        alert("Gagal memuat library PDF. Periksa koneksi internet.");
        return;
    }
    
    const reader = new FileReader();
    reader.onload = async (event) => {
        try {
            const arrayBuffer = event.target.result;
            const pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            const page = await pdfDoc.getPage(1);
            const viewport = page.getViewport({ scale: 3.0 });
            
            const tempCanvas = document.createElement('canvas');
            const tempCtx = tempCanvas.getContext('2d');
            tempCanvas.width = viewport.width;
            tempCanvas.height = viewport.height;
            
            await page.render({ canvasContext: tempCtx, viewport: viewport }).promise;
            
            const img = new Image();
            img.onload = () => {
                setBackgroundImage(img);
            };
            img.src = tempCanvas.toDataURL('image/png');
            
        } catch (error) {
            alert(`Gagal membaca PDF: ${error.message}`);
            uploadInput.value = '';
        }
    };
    reader.readAsArrayBuffer(file);
}

function setBackgroundImage(img) {
    state.backgroundImage = img;
    canvas.width = img.width;
    canvas.height = img.height;
    
    if (state.textFields.length === 0) {
        addNewTextField(state.dataHeaders.length > 0 ? state.dataHeaders[0] : null);
    }
    
    redrawCanvas();
    showEditor();
}

function showEditor() {
    editorSection.style.display = 'block';
    bulkSection.style.display = 'block';
    uploadSection.style.display = 'none';
    
    if (state.dataList.length > 0) {
        generateBulkBtn.disabled = false;
    }
}

// ============================================
// BULK DATA UPLOAD
// ============================================

function handleBulkFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    const ext = file.name.split('.').pop().toLowerCase();
    
    if (ext === 'csv') {
        reader.onload = (event) => processNamaList(event.target.result, true);
        reader.readAsText(file);
    } else if (['xls', 'xlsx'].includes(ext)) {
        reader.onload = (event) => parseExcelFile(event.target.result);
        reader.readAsArrayBuffer(file);
    } else {
        showStatus("Format file tidak didukung. Gunakan .csv, .xls, atau .xlsx", 'error');
    }
}

function parseExcelFile(arrayBuffer) {
    try {
        if (typeof XLSX === 'undefined') {
            throw new Error("Library SheetJS tidak termuat.");
        }
        
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const headers = XLSX.utils.sheet_to_json(worksheet, { header: 1 })[0];
        const json = XLSX.utils.sheet_to_json(worksheet);
        
        processNamaList(json, false, headers);
    } catch (error) {
        showStatus(`Gagal membaca Excel: ${error.message}`, 'error');
        generateBulkBtn.disabled = true;
    }
}

async function fetchGoogleSheet() {
    const url = gdocLinkInput.value.trim();
    if (!url) {
        showStatus("Harap masukkan URL Google Sheet.", 'error');
        return;
    }
    
    showStatus("Mengambil data...", 'info');
    
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`Status: ${response.status}`);
        const csvContent = await response.text();
        processNamaList(csvContent, true);
    } catch (error) {
        showStatus(`Gagal: ${error.message}. Pastikan link dipublikasikan sebagai CSV.`, 'error');
    }
}

// ============================================
// TEMPLATE MANAGEMENT
// ============================================

function saveTemplate() {
    if (state.textFields.length === 0) {
        alert("Tidak ada bidang teks untuk disimpan.");
        return;
    }
    
    const dataStr = JSON.stringify(state.textFields, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement('a');
    link.download = 'sertifikat_template.json';
    link.href = url;
    link.click();
    
    URL.revokeObjectURL(url);
}

function loadTemplate(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    if (!state.backgroundImage) {
        alert("Harap unggah desain sertifikat terlebih dahulu.");
        loadTemplateInput.value = '';
        return;
    }
    
    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const parsed = JSON.parse(event.target.result);
            if (!Array.isArray(parsed)) throw new Error("File template tidak valid.");
            
            state.textFields = parsed;
            state.nextTextId = state.textFields.length > 0 
                ? Math.max(...state.textFields.map(f => f.id)) + 1 
                : 0;
                
            redrawCanvas();
            selectTextField(null);
            alert("Template berhasil dimuat!");
        } catch (error) {
            alert(`Gagal memuat template: ${error.message}`);
        } finally {
            loadTemplateInput.value = '';
        }
    };
    reader.readAsText(file);
}

// ============================================
// BULK GENERATION
// ============================================

async function generateBulkCertificates() {
    if (state.dataList.length === 0 || !state.backgroundImage?.src) {
        alert("Mohon siapkan desain dan daftar nama terlebih dahulu.");
        return;
    }
    
    generateBulkBtn.disabled = true;
    showStatus("Memulai generasi massal...", 'info');
    
    const previouslySelected = state.selectedTextId;
    selectTextField(null);
    
    // Small delay to allow UI to update
    await new Promise(resolve => setTimeout(resolve, 50));
    
    try {
        const format = bulkFormatSelect.value;
        const asZip = zipCheckbox.checked;
        
        let jsPDF, zip;
        if (format === 'pdf') jsPDF = window.jspdf.jsPDF;
        if (asZip) zip = new JSZip();
        
        let generatedCount = 0;
        
        for (const row of state.dataList) {
            await generateSingleCertificate(row, format, asZip, zip, jsPDF);
            generatedCount++;
            
            const statusMsg = asZip 
                ? `Memproses ${generatedCount}/${state.dataList.length} ke ZIP...`
                : `Mengunduh ${generatedCount}/${state.dataList.length}...`;
            showStatus(statusMsg, 'info');
            
            // Small delay to prevent browser freezing
            if (!asZip) await new Promise(r => setTimeout(r, 100));
        }
        
        if (asZip) {
            showStatus("Membuat file ZIP...", 'info');
            const content = await zip.generateAsync({ type: "blob" });
            downloadBlob(content, 'Sertifikat_Massal.zip');
        }
        
        showStatus(`Selesai! ${generatedCount} sertifikat telah diproses.`, 'success');
        
    } catch (error) {
        showStatus(`Gagal: ${error.message}`, 'error');
    } finally {
        generateBulkBtn.disabled = false;
        selectTextField(previouslySelected);
    }
}

async function generateSingleCertificate(row, format, asZip, zip, jsPDF) {
    // Clear and draw background
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(state.backgroundImage, 0, 0, canvas.width, canvas.height);
    
    // Draw all text fields
    state.textFields.forEach(field => {
        drawFieldForBulk(field, row);
    });
    
    // Generate filename
    const firstKey = state.dataHeaders[0] || 'sertifikat';
    const nameValue = row[firstKey] || `sertifikat_${Date.now()}`;
    const cleanName = toTitleCase(String(nameValue)).replace(/[\\/*?:"<>|]/g, '');
    const fileName = `S - ${cleanName}`;
    
    if (asZip) {
        await addToZip(format, zip, fileName, jsPDF);
    } else {
        downloadSingle(format, fileName, jsPDF);
    }
}

function drawFieldForBulk(field, row) {
    let style = "";
    if (field.isItalic) style += "italic ";
    if (field.isBold) style += "bold ";
    
    ctx.font = `${style}${field.size}px "${field.family}"`;
    ctx.fillStyle = field.color;
    ctx.textAlign = field.align;
    ctx.textBaseline = 'middle';
    
    let text = field.dataLink ? (row[field.dataLink] || '') : field.text;
    
    // Apply text transformation
    if (field.dataLink) {
        if (field.transform === 'titlecase') text = toTitleCase(text);
        else if (field.transform === 'uppercase') text = String(text).toUpperCase();
    }
    
    ctx.fillText(text, field.x, field.y);
}

async function addToZip(format, zip, fileName, jsPDF) {
    if (format === 'png') {
        const dataUrl = canvas.toDataURL('image/png');
        const base64 = dataUrl.replace(/^data:image\/png;base64,/, "");
        zip.file(`${fileName}.png`, base64, { base64: true });
    } else {
        const pdf = createPdfFromCanvas(jsPDF);
        const blob = pdf.output('blob');
        zip.file(`${fileName}.pdf`, blob);
    }
}

function downloadSingle(format, fileName, jsPDF) {
    if (format === 'png') {
        const link = document.createElement('a');
        link.download = `${fileName}.png`;
        link.href = canvas.toDataURL('image/png');
        link.click();
    } else {
        const pdf = createPdfFromCanvas(jsPDF);
        pdf.save(`${fileName}.pdf`);
    }
}

function createPdfFromCanvas(jsPDF) {
    const orientation = canvas.width > canvas.height ? 'l' : 'p';
    const doc = new jsPDF({
        orientation: orientation,
        unit: 'px',
        format: [canvas.width, canvas.height]
    });
    doc.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, canvas.width, canvas.height);
    return doc;
}

function downloadBlob(blob, filename) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

function setupEyedropper() {
    if (!('EyeDropper' in window)) {
        colorPickerBtn.style.display = 'none';
        return;
    }
    
    colorPickerBtn.addEventListener('click', async () => {
        try {
            const eyeDropper = new EyeDropper();
            const result = await eyeDropper.open();
            fontColorInput.value = result.sRGBHex;
            fontColorHexInput.value = result.sRGBHex;
            updateSelectedTextField('color', result.sRGBHex);
        } catch (e) {
            console.log("Eyedropper dibatalkan.");
        }
    });
}

function handleAddFont() {
    const name = newFontNameInput.value.trim();
    const file = newFontFileInput.files[0];
    
    if (!name || !file) {
        alert("Harap isi Nama Font dan pilih File Font.");
        return;
    }
    
    const reader = new FileReader();
    reader.onload = (event) => {
        const font = new FontFace(name, `url(${event.target.result})`);
        font.load().then(loaded => {
            document.fonts.add(loaded);
            
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            fontFamilySelect.appendChild(option);
            fontFamilySelect.value = name;
            
            if (state.selectedTextId !== null) {
                updateSelectedTextField('family', name);
            }
            
            newFontNameInput.value = '';
            newFontFileInput.value = '';
            alert(`Font '${name}' berhasil ditambahkan!`);
        }).catch(err => {
            alert(`Gagal memuat font: ${err.message}`);
        });
    };
    reader.readAsDataURL(file);
}

function getMousePos(e) {
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    return {
        x: (e.clientX - rect.left) * scaleX,
        y: (e.clientY - rect.top) * scaleY
    };
}

// ============================================
// EVENT LISTENERS SETUP
// ============================================

function initEventListeners() {
    // File uploads
    uploadInput.addEventListener('change', handleDesignUpload);
    bulkFileInput.addEventListener('change', handleBulkFileUpload);
    fetchGdocBtn.addEventListener('click', fetchGoogleSheet);
    
    // Text management
    addTextBtn.addEventListener('click', () => addNewTextField());
    deleteTextBtn.addEventListener('click', () => {
        if (confirm("Hapus bidang teks ini?")) {
            state.textFields = state.textFields.filter(t => t.id !== state.selectedTextId);
            selectTextField(null);
        }
    });
    
    // Form inputs
    dataLinkSelect.addEventListener('change', () => updateSelectedTextField('dataLink', dataLinkSelect.value));
    textInput.addEventListener('input', () => updateSelectedTextField('text', textInput.value));
    fontFamilySelect.addEventListener('change', () => updateSelectedTextField('family', fontFamilySelect.value));
    fontSizeInput.addEventListener('change', () => updateSelectedTextField('size', parseInt(fontSizeInput.value)));
    fontAlignSelect.addEventListener('change', () => updateSelectedTextField('align', fontAlignSelect.value));
    textTransformSelect.addEventListener('change', () => updateSelectedTextField('transform', textTransformSelect.value));
    
    // Color inputs
    fontColorInput.addEventListener('input', () => {
        fontColorHexInput.value = fontColorInput.value;
        updateSelectedTextField('color', fontColorInput.value);
    });
    fontColorHexInput.addEventListener('change', () => {
        let val = fontColorHexInput.value;
        if (val.length === 6 && !val.startsWith('#')) val = '#' + val;
        if (/^#[0-9A-F]{6}$/i.test(val)) {
            fontColorInput.value = val;
            updateSelectedTextField('color', val);
        } else {
            fontColorHexInput.value = fontColorInput.value;
        }
    });
    
    // Style buttons
    fontBoldBtn.addEventListener('click', () => {
        const field = state.textFields.find(t => t.id === state.selectedTextId);
        if (field) {
            field.isBold = !field.isBold;
            fontBoldBtn.classList.toggle('active', field.isBold);
            redrawCanvas();
        }
    });
    
    fontItalicBtn.addEventListener('click', () => {
        const field = state.textFields.find(t => t.id === state.selectedTextId);
        if (field) {
            field.isItalic = !field.isItalic;
            fontItalicBtn.classList.toggle('active', field.isItalic);
            redrawCanvas();
        }
    });
    
    // Position nudge buttons
    moveUpBtn.addEventListener('click', () => nudgeField(0, -NUDGE_AMOUNT));
    moveDownBtn.addEventListener('click', () => nudgeField(0, NUDGE_AMOUNT));
    moveLeftBtn.addEventListener('click', () => nudgeField(-NUDGE_AMOUNT, 0));
    moveRightBtn.addEventListener('click', () => nudgeField(NUDGE_AMOUNT, 0));
    
    // Canvas interactions
    canvas.addEventListener('mousedown', handleCanvasMouseDown);
    canvas.addEventListener('mousemove', handleCanvasMouseMove);
    canvas.addEventListener('mouseup', handleCanvasMouseUp);
    canvas.addEventListener('mouseout', handleCanvasMouseUp);
    
    // Template & download
    saveTemplateBtn.addEventListener('click', saveTemplate);
    loadTemplateBtn.addEventListener('click', () => loadTemplateInput.click());
    loadTemplateInput.addEventListener('change', loadTemplate);
    downloadBtn.addEventListener('click', downloadSingleCertificate);
    generateBulkBtn.addEventListener('click', generateBulkCertificates);
    
    // Font management
    addFontBtn.addEventListener('click', handleAddFont);
}

function nudgeField(dx, dy) {
    const field = state.textFields.find(t => t.id === state.selectedTextId);
    if (field) {
        field.x += dx;
        field.y += dy;
        redrawCanvas();
    }
}

// Canvas interaction handlers
function handleCanvasMouseDown(e) {
    if (!state.backgroundImage?.src) return;
    
    const pos = getMousePos(e);
    const clickedField = findFieldAtPosition(pos);
    
    if (clickedField) {
        state.isDragging = true;
        selectTextField(clickedField.id);
        canvas.style.cursor = 'move';
        state.dragStartOffset = {
            x: pos.x - clickedField.x,
            y: pos.y - clickedField.y
        };
    } else {
        state.isDragging = false;
        selectTextField(null);
        canvas.style.cursor = 'crosshair';
    }
}

function handleCanvasMouseMove(e) {
    if (!state.isDragging || state.selectedTextId === null) return;
    
    const field = state.textFields.find(t => t.id === state.selectedTextId);
    if (!field) return;
    
    const pos = getMousePos(e);
    field.x = pos.x - state.dragStartOffset.x;
    field.y = pos.y - state.dragStartOffset.y;
    redrawCanvas();
}

function handleCanvasMouseUp() {
    state.isDragging = false;
    canvas.style.cursor = 'crosshair';
}

function findFieldAtPosition(pos) {
    for (let i = state.textFields.length - 1; i >= 0; i--) {
        const field = state.textFields[i];
        const box = field.boundingBox;
        if (pos.x >= box.x && pos.x <= box.x + box.width &&
            pos.y >= box.y && pos.y <= box.y + box.height) {
            return field;
        }
    }
    return null;
}

function downloadSingleCertificate() {
    if (!state.backgroundImage?.src) {
        alert("Mohon upload desain sertifikat terlebih dahulu.");
        return;
    }
    
    const prevSelected = state.selectedTextId;
    selectTextField(null);
    redrawCanvas();
    
    const link = document.createElement('a');
    link.download = 'Sertifikat_Preview.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
    
    selectTextField(prevSelected);
}

// ============================================
// INITIALIZATION
// ============================================

document.addEventListener('DOMContentLoaded', () => {
    initSplitPane();
    initEventListeners();
    setupEyedropper();
    updateDataLinkDropdown();
    
    // Hide editor sections initially
    editorSection.style.display = 'none';
    bulkSection.style.display = 'none';
});
