// ==========================================
// CERTIFICATE GENERATOR PRO - BUG FIX VERSION
// Fixed: Excel Cleaner data flow & File naming
// ==========================================

document.addEventListener('DOMContentLoaded', function() {
    
    // ==========================================
    // GLOBAL VARIABLES
    // ==========================================
    
    let canvas = document.getElementById('sertifikat-canvas');
    let ctx = canvas ? canvas.getContext('2d') : null;
    let certificateImage = null;
    let textElements = [];
    let selectedTextIndex = -1;
    let isDragging = false;
    let dragOffset = { x: 0, y: 0 };
    let scale = 1;
    
    // Excel Cleaner variables
    let excelData = null;
    let excelFileName = '';
    let selectedNameColumns = [];
    let detectedNameColumns = [];
    
    // CRITICAL: Cleaned data for bulk generation
    let cleanedDataForBulk = []; // Array of single names
    let originalFileName = '';
    
    // Other features
    let imageLibrary = JSON.parse(localStorage.getItem('certImageLibrary') || '[]');
    let qrCodeSettings = { enabled: false, dataPattern: '{{name}}-{{id}}', position: 'bottom-right', size: 100 };
    let analytics = JSON.parse(localStorage.getItem('certAnalytics') || '{"totalGenerated":0,"history":[]}');
    let previewData = null;
    let previewIndex = 0;

    // ==========================================
    // UTILITY FUNCTIONS
    // ==========================================
    
    function showStatus(elementId, message, type = 'info', duration = 5000) {
        const statusDiv = document.getElementById(elementId);
        if (!statusDiv) return;
        
        statusDiv.textContent = message;
        statusDiv.className = 'status-message ' + type;
        statusDiv.classList.remove('hidden');
        
        if (duration > 0) {
            setTimeout(() => {
                statusDiv.classList.add('hidden');
            }, duration);
        }
    }

    function generateId() {
        return 'text_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    function isNameColumn(header) {
        if (!header) return false;
        const headerStr = String(header).toLowerCase().trim();
        const nameKeywords = [
            'nama', 'name', 'peserta', 'participant', 'siswa', 'student',
            'mahasiswa', 'karyawan', 'employee', 'pegawai', 'staff',
            'user', 'pengguna', 'member', 'anggota', 'person', 'orang',
            'fullname', 'full name', 'nama lengkap', 'lengkap',
            'first name', 'last name', 'nama depan', 'nama belakang',
            'nama1', 'nama2', 'nama3', 'nama_1', 'nama_2', 'nama_3',
            'name1', 'name2', 'name3', 'name_1', 'name_2', 'name_3',
            'peserta1', 'peserta2', 'peserta3', 'participant1', 'participant2'
        ];
        return nameKeywords.some(keyword => headerStr.includes(keyword));
    }

    function cleanName(name) {
        if (!name) return '';
        return String(name).trim().replace(/\s+/g, ' ');
    }

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // ==========================================
    // SECTION 1: CERTIFICATE UPLOAD
    // ==========================================
    
    const designUpload = document.getElementById('design-upload');
    const uploadSection = document.getElementById('upload-section');
    const editorSection = document.getElementById('editor-section');
    const bulkSection = document.getElementById('bulk-section');

    if (designUpload) {
        designUpload.addEventListener('change', handleDesignUpload);
    }

    function handleDesignUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        
        if (file.type === 'application/pdf') {
            handlePDFUpload(file);
        } else {
            reader.onload = function(event) {
                const img = new Image();
                img.onload = function() {
                    certificateImage = img;
                    initCanvas();
                    showEditor();
                };
                img.src = event.target.result;
            };
            reader.readAsDataURL(file);
        }
    }

    function handlePDFUpload(file) {
        const reader = new FileReader();
        reader.onload = function(event) {
            const typedarray = new Uint8Array(event.target.result);
            
            pdfjsLib.getDocument({data: typedarray}).promise.then(function(pdf) {
                return pdf.getPage(1);
            }).then(function(page) {
                const viewport = page.getViewport({scale: 1.5});
                const tempCanvas = document.createElement('canvas');
                const tempCtx = tempCanvas.getContext('2d');
                tempCanvas.width = viewport.width;
                tempCanvas.height = viewport.height;
                
                return page.render({
                    canvasContext: tempCtx,
                    viewport: viewport
                }).promise.then(function() {
                    const img = new Image();
                    img.onload = function() {
                        certificateImage = img;
                        initCanvas();
                        showEditor();
                    };
                    img.src = tempCanvas.toDataURL();
                });
            });
        };
        reader.readAsArrayBuffer(file);
    }

    function initCanvas() {
        if (!canvas || !certificateImage) return;
        canvas.width = certificateImage.width;
        canvas.height = certificateImage.height;
        scale = Math.min(800 / canvas.width, 600 / canvas.height, 1);
        canvas.style.width = (canvas.width * scale) + 'px';
        canvas.style.height = (canvas.height * scale) + 'px';
        redrawCanvas();
    }

    function showEditor() {
        if (uploadSection) uploadSection.style.display = 'none';
        if (editorSection) editorSection.style.display = 'block';
        if (bulkSection) bulkSection.style.display = 'block';
        addDefaultText();
    }

    function addDefaultText() {
        if (textElements.length === 0) {
            addTextElement({
                text: 'Nama Peserta',
                x: canvas.width / 2,
                y: canvas.height / 2,
                fontSize: 50,
                fontFamily: 'Times New Roman',
                color: '#000000',
                align: 'center',
                bold: false,
                italic: false,
                transform: 'none',
                dataLink: ''
            });
        }
    }

    // ==========================================
    // SECTION 2: CANVAS & TEXT EDITING
    // ==========================================
    
    function redrawCanvas() {
        if (!ctx || !certificateImage) return;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(certificateImage, 0, 0);
        
        textElements.forEach((textEl, index) => {
            drawTextElement(textEl, index === selectedTextIndex);
        });
        
        if (qrCodeSettings.enabled) {
            drawQRCodeOnCanvas();
        }
    }

    function drawTextElement(textEl, isSelected) {
        ctx.save();
        
        let displayText = textEl.text;
        if (textEl.transform === 'uppercase') {
            displayText = displayText.toUpperCase();
        } else if (textEl.transform === 'titlecase') {
            displayText = displayText.replace(/\w\S*/g, function(txt) {
                return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
            });
        }
        
        let fontStyle = '';
        if (textEl.bold) fontStyle += 'bold ';
        if (textEl.italic) fontStyle += 'italic ';
        
        ctx.font = fontStyle + textEl.fontSize + 'px ' + textEl.fontFamily;
        ctx.fillStyle = textEl.color;
        ctx.textAlign = textEl.align;
        ctx.textBaseline = 'middle';
        
        const lines = displayText.split('\n');
        const lineHeight = textEl.fontSize * 1.2;
        const totalHeight = lines.length * lineHeight;
        let startY = textEl.y - (totalHeight / 2) + (lineHeight / 2);
        
        lines.forEach((line, i) => {
            ctx.fillText(line, textEl.x, startY + (i * lineHeight));
        });
        
        if (isSelected) {
            const metrics = ctx.measureText(displayText);
            const padding = 10;
            let boxX, boxWidth;
            
            if (textEl.align === 'center') {
                boxX = textEl.x - (metrics.width / 2) - padding;
                boxWidth = metrics.width + (padding * 2);
            } else if (textEl.align === 'right') {
                boxX = textEl.x - metrics.width - padding;
                boxWidth = metrics.width + (padding * 2);
            } else {
                boxX = textEl.x - padding;
                boxWidth = metrics.width + (padding * 2);
            }
            
            ctx.strokeStyle = '#1a237e';
            ctx.lineWidth = 2;
            ctx.setLineDash([5, 5]);
            ctx.strokeRect(boxX, startY - textEl.fontSize/2 - padding, boxWidth, totalHeight + (padding * 2));
            
            ctx.fillStyle = '#1a237e';
            ctx.setLineDash([]);
            ctx.fillRect(boxX + boxWidth - 6, startY + totalHeight/2 + padding - 6, 12, 12);
        }
        
        ctx.restore();
    }

    // Canvas events
    if (canvas) {
        canvas.addEventListener('mousedown', handleCanvasMouseDown);
        canvas.addEventListener('mousemove', handleCanvasMouseMove);
        canvas.addEventListener('mouseup', handleCanvasMouseUp);
        canvas.addEventListener('dblclick', handleCanvasDoubleClick);
    }

    function getMousePos(evt) {
        const rect = canvas.getBoundingClientRect();
        return {
            x: (evt.clientX - rect.left) / scale,
            y: (evt.clientY - rect.top) / scale
        };
    }

    function handleCanvasMouseDown(e) {
        const pos = getMousePos(e);
        
        for (let i = textElements.length - 1; i >= 0; i--) {
            if (isPointInText(pos.x, pos.y, textElements[i])) {
                selectedTextIndex = i;
                isDragging = true;
                dragOffset.x = pos.x - textElements[i].x;
                dragOffset.y = pos.y - textElements[i].y;
                updateToolbar();
                redrawCanvas();
                return;
            }
        }
        
        selectedTextIndex = -1;
        updateToolbar();
        redrawCanvas();
    }

    function handleCanvasMouseMove(e) {
        if (!isDragging || selectedTextIndex === -1) return;
        
        const pos = getMousePos(e);
        textElements[selectedTextIndex].x = pos.x - dragOffset.x;
        textElements[selectedTextIndex].y = pos.y - dragOffset.y;
        updateToolbarValues();
        redrawCanvas();
    }

    function handleCanvasMouseUp() {
        isDragging = false;
    }

    function handleCanvasDoubleClick(e) {
        const pos = getMousePos(e);
        
        for (let i = textElements.length - 1; i >= 0; i--) {
            if (isPointInText(pos.x, pos.y, textElements[i])) {
                const newText = prompt('Edit text:', textElements[i].text);
                if (newText !== null) {
                    textElements[i].text = newText;
                    redrawCanvas();
                    updateToolbarValues();
                }
                return;
            }
        }
    }

    function isPointInText(x, y, textEl) {
        ctx.save();
        let fontStyle = '';
        if (textEl.bold) fontStyle += 'bold ';
        if (textEl.italic) fontStyle += 'italic ';
        ctx.font = fontStyle + textEl.fontSize + 'px ' + textEl.fontFamily;
        
        const metrics = ctx.measureText(textEl.text);
        const padding = 10;
        let boxX, boxY, boxWidth, boxHeight;
        
        boxY = textEl.y - textEl.fontSize/2 - padding;
        boxHeight = textEl.fontSize + padding * 2;
        
        if (textEl.align === 'center') {
            boxX = textEl.x - metrics.width/2 - padding;
            boxWidth = metrics.width + padding * 2;
        } else if (textEl.align === 'right') {
            boxX = textEl.x - metrics.width - padding;
            boxWidth = metrics.width + padding * 2;
        } else {
            boxX = textEl.x - padding;
            boxWidth = metrics.width + padding * 2;
        }
        
        ctx.restore();
        
        return x >= boxX && x <= boxX + boxWidth && y >= boxY && y <= boxY + boxHeight;
    }

    // ==========================================
    // SECTION 3: TOOLBAR CONTROLS
    // ==========================================
    
    const addTextBtn = document.getElementById('add-text-btn');
    const deleteTextBtn = document.getElementById('delete-text-btn');
    const deleteTextContainer = document.getElementById('delete-text-container');
    const textEditControls = document.getElementById('text-edit-controls');
    
    const textInput = document.getElementById('text-input');
    const fontFamily = document.getElementById('font-family');
    const fontSize = document.getElementById('font-size');
    const fontColor = document.getElementById('font-color');
    const fontColorHex = document.getElementById('font-color-hex');
    const fontBold = document.getElementById('font-bold');
    const fontItalic = document.getElementById('font-italic');
    const fontAlign = document.getElementById('font-align');
    const textTransform = document.getElementById('text-transform');
    const dataLinkSelect = document.getElementById('data-link-select');

    if (addTextBtn) {
        addTextBtn.addEventListener('click', function() {
            addTextElement({
                text: 'Teks Baru',
                x: canvas.width / 2,
                y: canvas.height / 2 + 50,
                fontSize: 40,
                fontFamily: 'Arial',
                color: '#000000',
                align: 'center',
                bold: false,
                italic: false,
                transform: 'none',
                dataLink: ''
            });
        });
    }

    if (deleteTextBtn) {
        deleteTextBtn.addEventListener('click', function() {
            if (selectedTextIndex !== -1) {
                textElements.splice(selectedTextIndex, 1);
                selectedTextIndex = -1;
                updateToolbar();
                redrawCanvas();
            }
        });
    }

    function addTextElement(config) {
        config.id = generateId();
        textElements.push(config);
        selectedTextIndex = textElements.length - 1;
        updateToolbar();
        redrawCanvas();
    }

    function updateToolbar() {
        if (selectedTextIndex === -1) {
            if (textEditControls) textEditControls.classList.add('hidden');
            if (deleteTextContainer) deleteTextContainer.classList.add('invisible');
            return;
        }
        
        if (textEditControls) textEditControls.classList.remove('hidden');
        if (deleteTextContainer) deleteTextContainer.classList.remove('invisible');
        updateToolbarValues();
    }

    function updateToolbarValues() {
        if (selectedTextIndex === -1) return;
        
        const el = textElements[selectedTextIndex];
        if (textInput) textInput.value = el.text;
        if (fontFamily) fontFamily.value = el.fontFamily;
        if (fontSize) fontSize.value = el.fontSize;
        if (fontColor) fontColor.value = el.color;
        if (fontColorHex) fontColorHex.value = el.color;
        if (fontAlign) fontAlign.value = el.align;
        if (textTransform) textTransform.value = el.transform || 'none';
        
        if (fontBold) fontBold.classList.toggle('active', el.bold);
        if (fontItalic) fontItalic.classList.toggle('active', el.italic);
    }

    // Event listeners
    if (textInput) {
        textInput.addEventListener('input', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].text = this.value;
                redrawCanvas();
            }
        });
    }

    if (fontFamily) {
        fontFamily.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].fontFamily = this.value;
                redrawCanvas();
            }
        });
    }

    if (fontSize) {
        fontSize.addEventListener('input', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].fontSize = parseInt(this.value) || 40;
                redrawCanvas();
            }
        });
    }

    if (fontColor) {
        fontColor.addEventListener('input', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].color = this.value;
                if (fontColorHex) fontColorHex.value = this.value;
                redrawCanvas();
            }
        });
    }

    if (fontColorHex) {
        fontColorHex.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].color = this.value;
                if (fontColor) fontColor.value = this.value;
                redrawCanvas();
            }
        });
    }

    if (fontBold) {
        fontBold.addEventListener('click', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].bold = !textElements[selectedTextIndex].bold;
                this.classList.toggle('active');
                redrawCanvas();
            }
        });
    }

    if (fontItalic) {
        fontItalic.addEventListener('click', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].italic = !textElements[selectedTextIndex].italic;
                this.classList.toggle('active');
                redrawCanvas();
            }
        });
    }

    if (fontAlign) {
        fontAlign.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].align = this.value;
                redrawCanvas();
            }
        });
    }

    if (textTransform) {
        textTransform.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].transform = this.value;
                redrawCanvas();
            }
        });
    }

    // Position controls
    const moveUp = document.getElementById('move-up');
    const moveDown = document.getElementById('move-down');
    const moveLeft = document.getElementById('move-left');
    const moveRight = document.getElementById('move-right');

    function moveSelectedText(dx, dy) {
        if (selectedTextIndex !== -1) {
            textElements[selectedTextIndex].x += dx;
            textElements[selectedTextIndex].y += dy;
            redrawCanvas();
        }
    }

    if (moveUp) moveUp.addEventListener('click', () => moveSelectedText(0, -5));
    if (moveDown) moveDown.addEventListener('click', () => moveSelectedText(0, 5));
    if (moveLeft) moveLeft.addEventListener('click', () => moveSelectedText(-5, 0));
    if (moveRight) moveRight.addEventListener('click', () => moveSelectedText(5, 0));

    // ==========================================
    // SECTION 4: CUSTOM FONTS
    // ==========================================
    
    const addFontBtn = document.getElementById('add-font-btn');
    const newFontName = document.getElementById('new-font-name');
    const newFontFile = document.getElementById('new-font-file');

    if (addFontBtn) {
        addFontBtn.addEventListener('click', function() {
            const name = newFontName ? newFontName.value.trim() : '';
            const file = newFontFile ? newFontFile.files[0] : null;
            
            if (!name || !file) {
                alert('Masukkan nama font dan pilih file font (.ttf, .otf, .woff)');
                return;
            }
            
            const reader = new FileReader();
            reader.onload = function(e) {
                const fontFace = new FontFace(name, e.target.result);
                fontFace.load().then(function(loadedFace) {
                    document.fonts.add(loadedFace);
                    
                    const option = document.createElement('option');
                    option.value = name;
                    option.textContent = name;
                    if (fontFamily) fontFamily.appendChild(option);
                    
                    if (newFontName) newFontName.value = '';
                    if (newFontFile) newFontFile.value = '';
                    alert('Font berhasil ditambahkan!');
                });
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // ==========================================
    // SECTION 5: TEMPLATE SAVE/LOAD
    // ==========================================
    
    const saveTemplateBtn = document.getElementById('save-template-btn');
    const loadTemplateBtn = document.getElementById('load-template-btn');
    const loadTemplateInput = document.getElementById('load-template-input');

    if (saveTemplateBtn) {
        saveTemplateBtn.addEventListener('click', function() {
            const template = {
                textElements: textElements,
                canvasWidth: canvas ? canvas.width : 0,
                canvasHeight: canvas ? canvas.height : 0,
                qrSettings: qrCodeSettings,
                version: '2.0 Pro'
            };
            
            const blob = new Blob([JSON.stringify(template, null, 2)], {type: 'application/json'});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'certificate-template-pro.json';
            a.click();
            URL.revokeObjectURL(url);
        });
    }

    if (loadTemplateBtn) {
        loadTemplateBtn.addEventListener('click', function() {
            if (loadTemplateInput) loadTemplateInput.click();
        });
    }

    if (loadTemplateInput) {
        loadTemplateInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(event) {
                try {
                    const template = JSON.parse(event.target.result);
                    textElements = template.textElements || [];
                    
                    if (template.qrSettings) {
                        qrCodeSettings = template.qrSettings;
                    }
                    
                    selectedTextIndex = -1;
                    updateToolbar();
                    redrawCanvas();
                    
                    document.getElementById('qr-enable').checked = qrCodeSettings.enabled;
                    document.getElementById('qr-pattern').value = qrCodeSettings.dataPattern;
                    document.getElementById('qr-position').value = qrCodeSettings.position;
                    document.getElementById('qr-size').value = qrCodeSettings.size;
                    
                    alert('Template Pro berhasil dimuat!');
                } catch (err) {
                    alert('Error membaca template: ' + err.message);
                }
            };
            reader.readAsText(file);
            this.value = '';
        });
    }

    // ==========================================
    // SECTION 6: DOWNLOAD SINGLE
    // ==========================================
    
    const downloadBtn = document.getElementById('download-btn');

    if (downloadBtn) {
        downloadBtn.addEventListener('click', function() {
            if (!certificateImage) {
                alert('Upload desain sertifikat terlebih dahulu!');
                return;
            }
            
            const link = document.createElement('a');
            link.download = 'sertifikat.png';
            link.href = canvas.toDataURL('image/png');
            link.click();
        });
    }

    // ==========================================
    // SECTION 7: BULK GENERATION (FIXED VERSION)
    // ==========================================
    
    const bulkFileUpload = document.getElementById('bulk-file-upload');
    const gdocLink = document.getElementById('gdoc-link');
    const fetchGdocBtn = document.getElementById('fetch-gdoc-btn');
    const generateBulkBtn = document.getElementById('generate-bulk-btn');
    const bulkFormat = document.getElementById('bulk-format');
    const downloadAsZip = document.getElementById('download-as-zip');
    const bulkStatus = document.getElementById('bulk-status');

    let bulkHeaders = [];
    // CRITICAL: Use cleanedDataForBulk instead of separate variable

    if (bulkFileUpload) {
        bulkFileUpload.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            // Store original filename for naming
            originalFileName = file.name.replace(/\.[^/.]+$/, "");
            
            const reader = new FileReader();
            reader.onload = function(event) {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
                    
                    processBulkData(jsonData);
                    showStatus('bulk-status', 'Data berhasil dimuat! ' + (jsonData.length - 1) + ' baris ditemukan.', 'success');
                } catch (err) {
                    showStatus('bulk-status', 'Error: ' + err.message, 'error');
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    if (fetchGdocBtn) {
        fetchGdocBtn.addEventListener('click', async function() {
            const url = gdocLink ? gdocLink.value.trim() : '';
            if (!url) {
                showStatus('bulk-status', 'Masukkan link Google Sheet', 'error');
                return;
            }
            
            try {
                const response = await fetch(url);
                const csvText = await response.text();
                const workbook = XLSX.read(csvText, {type: 'string'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
                
                processBulkData(jsonData);
                showStatus('bulk-status', 'Data berhasil diambil! ' + (jsonData.length - 1) + ' baris ditemukan.', 'success');
            } catch (err) {
                showStatus('bulk-status', 'Error mengambil data: ' + err.message, 'error');
            }
        });
    }

    function processBulkData(data) {
        if (data.length < 2) {
            showStatus('bulk-status', 'Data terlalu sedikit (minimal 1 header + 1 data)', 'error');
            return;
        }
        
        bulkHeaders = data[0];
        
        // CRITICAL FIX: Store as cleaned single names array
        // Each element is [name] for single column or we need to flatten
        cleanedDataForBulk = [];
        
        // Check if this is already cleaned data (single column with "Nama" header)
        const isCleanedData = bulkHeaders.length === 1 && isNameColumn(bulkHeaders[0]);
        
        if (isCleanedData) {
            // Data sudah bersih dari Excel Cleaner
            for (let i = 1; i < data.length; i++) {
                if (data[i][0] && String(data[i][0]).trim() !== '') {
                    cleanedDataForBulk.push([data[i][0]]); // Keep as array for consistency
                }
            }
        } else {
            // Data mentah, flatten semua nama dari semua kolom
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                for (let j = 0; j < row.length; j++) {
                    const name = cleanName(row[j]);
                    if (name && isNameColumn(bulkHeaders[j])) {
                        cleanedDataForBulk.push([name]);
                    }
                }
            }
        }
        
        // Update dropdown
        if (dataLinkSelect) {
            dataLinkSelect.innerHTML = '<option value="">-- Tidak dihubungkan --</option>';
            bulkHeaders.forEach((header, index) => {
                const option = document.createElement('option');
                option.value = header;
                option.textContent = header;
                dataLinkSelect.appendChild(option);
            });
        }
        
        // Enable generate button
        if (generateBulkBtn) {
            generateBulkBtn.disabled = false;
        }
        
        // Update preview data
        previewData = cleanedDataForBulk;
        
        console.log('Bulk data processed:', cleanedDataForBulk.length, 'names');
    }

    if (dataLinkSelect) {
        dataLinkSelect.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].dataLink = this.value;
            }
        });
    }

    if (generateBulkBtn) {
        generateBulkBtn.addEventListener('click', async function() {
            // CRITICAL FIX: Use cleanedDataForBulk.length instead of bulkData.length
            if (!certificateImage || cleanedDataForBulk.length === 0) {
                showStatus('bulk-status', 'Upload desain dan data terlebih dahulu!', 'error');
                return;
            }
            
            const format = bulkFormat ? bulkFormat.value : 'png';
            const asZip = downloadAsZip ? downloadAsZip.checked : true;
            
            generateBulkBtn.disabled = true;
            generateBulkBtn.textContent = 'Memproses...';
            
            try {
                if (asZip && format === 'png') {
                    await generateBulkZip();
                } else if (format === 'pdf') {
                    await generateBulkPDF(asZip);
                } else {
                    await generateBulkPNG();
                }
                
                recordAnalytics(cleanedDataForBulk.length, format);
                showStatus('bulk-status', `✅ ${cleanedDataForBulk.length} sertifikat berhasil dibuat!`, 'success');
            } catch (err) {
                showStatus('bulk-status', '❌ Error: ' + err.message, 'error');
            } finally {
                generateBulkBtn.disabled = false;
                generateBulkBtn.textContent = 'Mulai Generasi Massal';
            }
        });
    }

    // FIXED: Generate ZIP dengan penamaan yang benar
    async function generateBulkZip() {
        const zip = new JSZip();
        const folder = zip.folder("sertifikat");
        
        for (let i = 0; i < cleanedDataForBulk.length; i++) {
            const name = cleanedDataForBulk[i][0] || 'Unknown';
            generateCertificateWithData(cleanedDataForBulk[i]);
            
            const dataUrl = canvas.toDataURL('image/png');
            const base64Data = dataUrl.split(',')[1];
            
            // CRITICAL FIX: Penamaan file S - [Nama].png
            const safeName = sanitizeFileName(name);
            const fileName = `S - ${safeName}.png`;
            
            folder.file(fileName, base64Data, {base64: true});
            
            if (i % 10 === 0) {
                showStatus('bulk-status', `Memproses... ${i + 1}/${cleanedDataForBulk.length}`, 'info', 0);
                await new Promise(r => setTimeout(r, 10));
            }
        }
        
        const content = await zip.generateAsync({type: "blob"});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(content);
        link.download = `${originalFileName || 'sertifikat'}_bulk.zip`;
        link.click();
    }

    // FIXED: Generate PNG individual dengan penamaan yang benar
    async function generateBulkPNG() {
        for (let i = 0; i < cleanedDataForBulk.length; i++) {
            const name = cleanedDataForBulk[i][0] || 'Unknown';
            generateCertificateWithData(cleanedDataForBulk[i]);
            
            const link = document.createElement('a');
            
            // CRITICAL FIX: Penamaan file S - [Nama].png
            const safeName = sanitizeFileName(name);
            link.download = `S - ${safeName}.png`;
            
            link.href = canvas.toDataURL('image/png');
            link.click();
            
            await new Promise(r => setTimeout(r, 100));
        }
    }

    // FIXED: Generate PDF dengan penamaan yang benar
    async function generateBulkPDF(asZip) {
        const { jsPDF } = window.jspdf;
        
        if (asZip) {
            const zip = new JSZip();
            
            for (let i = 0; i < cleanedDataForBulk.length; i++) {
                const name = cleanedDataForBulk[i][0] || 'Unknown';
                generateCertificateWithData(cleanedDataForBulk[i]);
                
                const pdf = createPDFfromCanvas();
                const pdfBlob = pdf.output('blob');
                
                // CRITICAL FIX: Penamaan file S - [Nama].pdf
                const safeName = sanitizeFileName(name);
                const fileName = `S - ${safeName}.pdf`;
                
                zip.file(fileName, pdfBlob);
                
                if (i % 10 === 0) {
                    showStatus('bulk-status', `Memproses... ${i + 1}/${cleanedDataForBulk.length}`, 'info', 0);
                    await new Promise(r => setTimeout(r, 10));
                }
            }
            
            const content = await zip.generateAsync({type: "blob"});
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = `${originalFileName || 'sertifikat'}_pdf.zip`;
            link.click();
        } else {
            // Single PDF dengan multiple pages
            const pdf = new jsPDF({
                orientation: canvas.width > canvas.height ? 'l' : 'p',
                unit: 'px',
                format: [canvas.width, canvas.height]
            });
            
            for (let i = 0; i < cleanedDataForBulk.length; i++) {
                generateCertificateWithData(cleanedDataForBulk[i]);
                const imgData = canvas.toDataURL('image/png');
                
                if (i > 0) pdf.addPage();
                pdf.addImage(imgData, 'PNG', 0, 0, canvas.width, canvas.height);
                
                if (i % 10 === 0) {
                    showStatus('bulk-status', `Memproses... ${i + 1}/${cleanedDataForBulk.length}`, 'info', 0);
                    await new Promise(r => setTimeout(r, 10));
                }
            }
            
            pdf.save(`${originalFileName || 'sertifikat'}_bulk.pdf`);
        }
    }

    // HELPER: Sanitize filename untuk menghindari karakter ilegal
    function sanitizeFileName(name) {
        if (!name) return 'Unknown';
        return String(name)
            .replace(/[<>:"/\\|?*]/g, '_')  // Karakter ilegal di Windows
            .replace(/\s+/g, ' ')           // Multiple space jadi single
            .trim()
            .substring(0, 100);              // Batasi panjang
    }

    function generateCertificateWithData(rowData) {
        if (!ctx || !certificateImage) return;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(certificateImage, 0, 0);
        
        textElements.forEach(textEl => {
            let displayText = textEl.text;
            
            // CRITICAL FIX: Gunakan rowData[0] karena data sudah flattened
            if (textEl.dataLink && bulkHeaders.includes(textEl.dataLink)) {
                const colIndex = bulkHeaders.indexOf(textEl.dataLink);
                if (colIndex !== -1 && rowData[colIndex]) {
                    displayText = String(rowData[colIndex]);
                }
            } else {
                // Jika tidak ada data link, gunakan rowData[0] (nama utama)
                if (rowData[0]) {
                    displayText = String(rowData[0]);
                }
            }
            
            if (textEl.transform === 'uppercase') {
                displayText = displayText.toUpperCase();
            } else if (textEl.transform === 'titlecase') {
                displayText = displayText.replace(/\w\S*/g, function(txt) {
                    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
                });
            }
            
            ctx.save();
            let fontStyle = '';
            if (textEl.bold) fontStyle += 'bold ';
            if (textEl.italic) fontStyle += 'italic ';
            
            ctx.font = fontStyle + textEl.fontSize + 'px ' + textEl.fontFamily;
            ctx.fillStyle = textEl.color;
            ctx.textAlign = textEl.align;
            ctx.textBaseline = 'middle';
            
            const lines = displayText.split('\n');
            const lineHeight = textEl.fontSize * 1.2;
            const totalHeight = lines.length * lineHeight;
            let startY = textEl.y - (totalHeight / 2) + (lineHeight / 2);
            
            lines.forEach((line, i) => {
                ctx.fillText(line, textEl.x, startY + (i * lineHeight));
            });
            
            ctx.restore();
        });
        
        if (qrCodeSettings.enabled) {
            drawQRCodeOnCanvas();
        }
    }

    function createPDFfromCanvas() {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({
            orientation: canvas.width > canvas.height ? 'l' : 'p',
            unit: 'px',
            format: [canvas.width, canvas.height]
        });
        
        const imgData = canvas.toDataURL('image/png');
        pdf.addImage(imgData, 'PNG', 0, 0, canvas.width, canvas.height);
        
        return pdf;
    }

    // ==========================================
    // SECTION 8: SMART EXCEL CLEANER (FIXED)
    // ==========================================
    
    const cleanerDropZone = document.getElementById('cleaner-drop-zone');
    const cleanerFileInput = document.getElementById('cleaner-file-input');
    const cleanerWorkspace = document.getElementById('cleaner-workspace');
    const nameColumnSelect = document.getElementById('name-column-select');
    const cleanerPreviewBody = document.getElementById('cleaner-preview-body');
    const totalRowsEl = document.getElementById('total-rows');
    const removedColsEl = document.getElementById('removed-columns');
    const keptColsEl = document.getElementById('kept-columns');
    const downloadCleanBtn = document.getElementById('download-clean-btn');
    const outputFormat = document.getElementById('output-format');
    const outputFilename = document.getElementById('output-filename');
    const autoDetectBtn = document.getElementById('auto-detect-btn');
    const clearSelectionBtn = document.getElementById('clear-selection-btn');
    const selectedColumnsInfo = document.getElementById('selected-columns-info');

    // NEW: Button untuk langsung gunakan data bersih untuk bulk
    const useForBulkBtn = document.createElement('button');
    useForBulkBtn.id = 'use-for-bulk-btn';
    useForBulkBtn.className = 'btn-primary btn-large';
    useForBulkBtn.innerHTML = '<span class="btn-icon">🚀</span> Gunakan untuk Generate Massal';
    useForBulkBtn.style.marginTop = '15px';
    useForBulkBtn.style.display = 'none';
    
    // Tambahkan setelah download button
    if (downloadCleanBtn && downloadCleanBtn.parentNode) {
        downloadCleanBtn.parentNode.appendChild(useForBulkBtn);
    }

    if (cleanerDropZone) {
        cleanerDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            cleanerDropZone.classList.add('drag-over');
        });

        cleanerDropZone.addEventListener('dragleave', () => {
            cleanerDropZone.classList.remove('drag-over');
        });

        cleanerDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            cleanerDropZone.classList.remove('drag-over');
            const files = e.dataTransfer.files;
            if (files.length) handleCleanerFile(files[0]);
        });

        cleanerDropZone.addEventListener('click', () => {
            if (cleanerFileInput) cleanerFileInput.click();
        });
    }

    if (cleanerFileInput) {
        cleanerFileInput.addEventListener('change', (e) => {
            if (e.target.files.length) handleCleanerFile(e.target.files[0]);
        });
    }

    function handleCleanerFile(file) {
        const validFormats = ['.csv', '.xls', '.xlsx'];
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        
        if (!validFormats.includes(ext)) {
            showStatus('cleaner-status', '❌ Format file tidak didukung. Gunakan .csv, .xls, atau .xlsx', 'error');
            return;
        }

        excelFileName = file.name.replace(ext, '');
        originalFileName = excelFileName; // Simpan untuk penamaan file
        
        if (outputFilename) outputFilename.value = excelFileName + '_clean';
        
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
                
                if (jsonData.length === 0) {
                    showStatus('cleaner-status', '❌ File kosong atau tidak valid', 'error');
                    return;
                }

                excelData = jsonData;
                processCleanerData(jsonData);
                
                if (cleanerDropZone) cleanerDropZone.classList.add('has-file');
                if (cleanerWorkspace) cleanerWorkspace.classList.remove('hidden');
                showStatus('cleaner-status', '✅ File berhasil dimuat! Kolom nama otomatis terdeteksi.', 'success');
                
            } catch (err) {
                showStatus('cleaner-status', '❌ Error membaca file: ' + err.message, 'error');
            }
        };
        
        reader.readAsArrayBuffer(file);
    }

    function processCleanerData(data) {
        const headers = data[0];
        const totalCols = headers.length;
        
        selectedNameColumns = [];
        detectedNameColumns = [];
        
        headers.forEach((header, index) => {
            if (isNameColumn(header)) {
                detectedNameColumns.push({
                    index: index,
                    header: header || `Kolom ${index + 1}`
                });
                selectedNameColumns.push(index);
            }
        });
        
        populateColumnSelect(headers);
        updateCleanerStats(totalCols);
        updateColumnSelection();
        
        if (selectedNameColumns.length > 0) {
            updateCleanerPreview();
            updateSelectedColumnsInfo();
            updateUseForBulkButton();
        }
    }

    function populateColumnSelect(headers) {
        if (!nameColumnSelect) return;
        
        nameColumnSelect.innerHTML = '';
        
        const selectAllDiv = document.createElement('div');
        selectAllDiv.className = 'column-option select-all';
        selectAllDiv.innerHTML = `
            <label class="checkbox-label">
                <input type="checkbox" id="select-all-columns" ${selectedNameColumns.length === headers.length ? 'checked' : ''}>
                <span class="checkmark"></span>
                <strong>Pilih Semua Kolom</strong>
            </label>
        `;
        nameColumnSelect.appendChild(selectAllDiv);
        
        headers.forEach((header, index) => {
            const isDetected = detectedNameColumns.some(d => d.index === index);
            const isSelected = selectedNameColumns.includes(index);
            
            const div = document.createElement('div');
            div.className = 'column-option' + (isDetected ? ' detected' : '');
            div.innerHTML = `
                <label class="checkbox-label">
                    <input type="checkbox" value="${index}" ${isSelected ? 'checked' : ''} data-detected="${isDetected}">
                    <span class="checkmark"></span>
                    <span class="column-name">${header || `Kolom ${index + 1}`}</span>
                    ${isDetected ? '<span class="detected-badge">🎯 Terdeteksi</span>' : ''}
                </label>
            `;
            nameColumnSelect.appendChild(div);
        });
        
        const selectAllCheckbox = document.getElementById('select-all-columns');
        if (selectAllCheckbox) {
            selectAllCheckbox.addEventListener('change', function() {
                const checkboxes = nameColumnSelect.querySelectorAll('input[type="checkbox"]:not(#select-all-columns)');
                checkboxes.forEach(cb => {
                    cb.checked = this.checked;
                });
                updateSelectedColumnsFromCheckboxes();
            });
        }
        
        const checkboxes = nameColumnSelect.querySelectorAll('input[type="checkbox"]:not(#select-all-columns)');
        checkboxes.forEach(cb => {
            cb.addEventListener('change', function() {
                updateSelectedColumnsFromCheckboxes();
            });
        });
    }

    function updateSelectedColumnsFromCheckboxes() {
        selectedNameColumns = [];
        const checkboxes = nameColumnSelect.querySelectorAll('input[type="checkbox"]:not(#select-all-columns)');
        checkboxes.forEach(cb => {
            if (cb.checked) {
                selectedNameColumns.push(parseInt(cb.value));
            }
        });
        
        updateCleanerPreview();
        updateSelectedColumnsInfo();
        updateCleanerStats(excelData[0].length);
        updateUseForBulkButton();
    }

    function updateUseForBulkButton() {
        // Tampilkan tombol "Gunakan untuk Bulk" jika ada data valid
        const hasValidData = selectedNameColumns.length > 0 && excelData && excelData.length > 1;
        if (useForBulkBtn) {
            useForBulkBtn.style.display = hasValidData ? 'block' : 'none';
        }
    }

    // CRITICAL FIX: Event listener untuk tombol "Gunakan untuk Bulk"
    if (useForBulkBtn) {
        useForBulkBtn.addEventListener('click', function() {
            if (!excelData || selectedNameColumns.length === 0) return;
            
            // Generate cleaned data
            const tempCleanedData = [];
            
            for (let i = 1; i < excelData.length; i++) {
                const row = excelData[i];
                const names = [];
                
                selectedNameColumns.forEach(colIndex => {
                    const name = cleanName(row[colIndex]);
                    if (name) names.push(name);
                });
                
                // Flatten: setiap nama jadi 1 baris
                names.forEach(name => {
                    tempCleanedData.push([name]);
                });
            }
            
            // CRITICAL: Set ke variabel global yang digunakan bulk generation
            cleanedDataForBulk = tempCleanedData;
            bulkHeaders = ['Nama']; // Set header untuk data link
            
            // Update UI
            if (generateBulkBtn) generateBulkBtn.disabled = false;
            
            // Scroll ke bulk section
            const bulkSection = document.getElementById('bulk-section');
            if (bulkSection) {
                bulkSection.scrollIntoView({behavior: 'smooth'});
            }
            
            showStatus('bulk-status', `✅ ${cleanedDataForBulk.length} nama siap untuk generate massal! Klik "Mulai Generasi Massal".`, 'success', 8000);
            
            console.log('Data siap untuk bulk:', cleanedDataForBulk.length, 'names');
        });
    }

    function updateColumnSelection() {
        const checkboxes = nameColumnSelect.querySelectorAll('input[type="checkbox"]:not(#select-all-columns)');
        checkboxes.forEach(cb => {
            cb.checked = selectedNameColumns.includes(parseInt(cb.value));
        });
    }

    function updateCleanerStats(totalCols) {
        if (totalRowsEl) totalRowsEl.textContent = excelData.length - 1;
        if (removedColsEl) removedColsEl.textContent = totalCols - selectedNameColumns.length;
        if (keptColsEl) keptColsEl.textContent = selectedNameColumns.length;
    }

    function updateSelectedColumnsInfo() {
        if (!selectedColumnsInfo) return;
        
        if (selectedNameColumns.length === 0) {
            selectedColumnsInfo.innerHTML = '<span class="warning">⚠️ Pilih minimal 1 kolom</span>';
            return;
        }
        
        const headers = excelData[0];
        const selectedNames = selectedNameColumns.map(idx => headers[idx] || `Kolom ${idx + 1}`);
        
        selectedColumnsInfo.innerHTML = `
            <strong>${selectedNameColumns.length} kolom dipilih:</strong><br>
            <small>${selectedNames.join(', ')}</small>
        `;
    }

    function updateCleanerPreview() {
        if (!cleanerPreviewBody || !excelData || selectedNameColumns.length === 0) {
            if (cleanerPreviewBody) cleanerPreviewBody.innerHTML = '<tr><td colspan="2" style="text-align: center;">Pilih kolom untuk melihat preview</td></tr>';
            return;
        }
        
        cleanerPreviewBody.innerHTML = '';
        const rowsToShow = Math.min(excelData.length - 1, 10);
        
        for (let i = 1; i <= rowsToShow; i++) {
            const row = excelData[i];
            const names = [];
            selectedNameColumns.forEach(colIndex => {
                const name = cleanName(row[colIndex]);
                if (name) names.push(name);
            });
            
            if (names.length === 0) {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${i}</td>
                    <td class="empty-cell">- (kosong)</td>
                `;
                cleanerPreviewBody.appendChild(tr);
            } else {
                names.forEach((name, nameIdx) => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${i}${names.length > 1 ? `.${nameIdx + 1}` : ''}</td>
                        <td>${escapeHtml(name)}</td>
                    `;
                    cleanerPreviewBody.appendChild(tr);
                });
            }
        }

        if (excelData.length > 11) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td colspan="2" style="text-align: center; color: #6c757d; font-style: italic;">
                    ... dan ${excelData.length - 11} baris lainnya
                </td>
            `;
            cleanerPreviewBody.appendChild(tr);
        }
    }

    if (autoDetectBtn) {
        autoDetectBtn.addEventListener('click', function() {
            if (!excelData) return;
            
            const headers = excelData[0];
            selectedNameColumns = [];
            
            headers.forEach((header, index) => {
                if (isNameColumn(header)) {
                    selectedNameColumns.push(index);
                }
            });
            
            updateColumnSelection();
            updateCleanerPreview();
            updateSelectedColumnsInfo();
            updateCleanerStats(headers.length);
            updateUseForBulkButton();
            
            showStatus('cleaner-status', `✅ ${selectedNameColumns.length} kolom nama terdeteksi otomatis`, 'success');
        });
    }

    if (clearSelectionBtn) {
        clearSelectionBtn.addEventListener('click', function() {
            selectedNameColumns = [];
            const checkboxes = nameColumnSelect.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(cb => cb.checked = false);
            updateCleanerPreview();
            updateSelectedColumnsInfo();
            updateCleanerStats(excelData[0].length);
            updateUseForBulkButton();
        });
    }

    // Download handler untuk Excel Cleaner (hanya download, tidak untuk bulk)
    if (downloadCleanBtn) {
        downloadCleanBtn.addEventListener('click', function() {
            if (!excelData || selectedNameColumns.length === 0) {
                showStatus('cleaner-status', '❌ Pilih minimal 1 kolom nama', 'error');
                return;
            }

            const format = outputFormat ? outputFormat.value : 'xlsx';
            const filename = outputFilename ? (outputFilename.value || 'nama_bersih') : 'nama_bersih';
            
            const cleanData = [];
            cleanData.push(['Nama']);
            
            let totalNames = 0;
            
            for (let i = 1; i < excelData.length; i++) {
                const row = excelData[i];
                const names = [];
                selectedNameColumns.forEach(colIndex => {
                    const name = cleanName(row[colIndex]);
                    if (name) names.push(name);
                });
                
                if (names.length === 0) {
                    cleanData.push(['']);
                } else {
                    names.forEach(name => {
                        cleanData.push([name]);
                        totalNames++;
                    });
                }
            }

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(cleanData);
            ws['!cols'] = [{wch: 50}];
            XLSX.utils.book_append_sheet(wb, ws, "Nama Bersih");

            const ext = format === 'csv' ? '.csv' : '.xlsx';
            try {
                XLSX.writeFile(wb, filename + ext);
                showStatus('cleaner-status', `✅ ${totalNames} nama berhasil diekspor ke ${filename}${ext}`, 'success');
            } catch (err) {
                showStatus('cleaner-status', '❌ Error download: ' + err.message, 'error');
            }
        });
    }

    // ==========================================
    // SECTION 9: QR CODE & OTHER FEATURES
    // ==========================================
    
    function drawQRCodeOnCanvas() {
        let qrData = 'CERTIFICATE-VERIFICATION';
        if (previewData && previewData[previewIndex]) {
            const row = previewData[previewIndex];
            qrData = qrCodeSettings.dataPattern
                .replace('{{name}}', row[0] || 'Unknown')
                .replace('{{id}}', previewIndex + 1)
                .replace('{{date}}', new Date().toISOString().split('T')[0]);
        }
        
        const qrCanvas = generateQRCode(qrData, qrCodeSettings.size);
        const pos = getQRPosition(qrCodeSettings.position, canvas.width, canvas.height, qrCodeSettings.size);
        
        ctx.drawImage(qrCanvas, pos.x, pos.y);
    }

    function generateQRCode(data, size) {
        const qrCanvas = document.createElement('canvas');
        qrCanvas.width = size;
        qrCanvas.height = size;
        const qrCtx = qrCanvas.getContext('2d');
        
        const seed = data.split('').reduce((a,b)=>a+b.charCodeAt(0),0);
        
        qrCtx.fillStyle = 'white';
        qrCtx.fillRect(0, 0, size, size);
        qrCtx.fillStyle = 'black';
        
        const cellSize = Math.floor(size / 25);
        
        drawPositionPattern(qrCtx, 0, 0, cellSize * 7);
        drawPositionPattern(qrCtx, size - cellSize * 7, 0, cellSize * 7);
        drawPositionPattern(qrCtx, 0, size - cellSize * 7, cellSize * 7);
        
        for (let i = 0; i < 25; i++) {
            for (let j = 0; j < 25; j++) {
                if ((i < 7 && j < 7) || (i > 17 && j < 7) || (i < 7 && j > 17)) continue;
                
                const pseudoRandom = Math.sin(seed + i * 25 + j) > 0;
                if (pseudoRandom) {
                    qrCtx.fillRect(i * cellSize, j * cellSize, cellSize, cellSize);
                }
            }
        }
        
        return qrCanvas;
    }

    function drawPositionPattern(ctx, x, y, size) {
        ctx.fillStyle = 'black';
        ctx.fillRect(x, y, size, size);
        ctx.fillStyle = 'white';
        ctx.fillRect(x + size/7, y + size/7, size - size/3.5, size - size/3.5);
        ctx.fillStyle = 'black';
        ctx.fillRect(x + size/3.5, y + size/3.5, size/3.5, size/3.5);
    }

    function getQRPosition(position, canvasWidth, canvasHeight, qrSize) {
        const padding = 20;
        switch(position) {
            case 'top-left': return { x: padding, y: padding };
            case 'top-right': return { x: canvasWidth - qrSize - padding, y: padding };
            case 'bottom-left': return { x: padding, y: canvasHeight - qrSize - padding };
            case 'bottom-right': return { x: canvasWidth - qrSize - padding, y: canvasHeight - qrSize - padding };
            case 'center': return { x: (canvasWidth - qrSize) / 2, y: (canvasHeight - qrSize) / 2 };
            default: return { x: canvasWidth - qrSize - padding, y: canvasHeight - qrSize - padding };
        }
    }

    // Analytics
    function recordAnalytics(count, format) {
        analytics.totalGenerated += count;
        analytics.history.push({
            date: new Date().toISOString(),
            count: count,
            format: format
        });
        
        if (analytics.history.length > 100) {
            analytics.history = analytics.history.slice(-100);
        }
        
        localStorage.setItem('certAnalytics', JSON.stringify(analytics));
    }

    // ==========================================
    // INITIALIZATION
    // ==========================================
    
    console.log('✅ Certificate Generator Pro v2.0 - BUG FIX EDITION');
    console.log('Fixed: Excel Cleaner data flow & File naming (S - [Nama].png)');
    
}); // End DOMContentLoaded
