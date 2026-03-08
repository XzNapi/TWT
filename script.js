// ==========================================
// CERTIFICATE GENERATOR + SMART EXCEL CLEANER + BONUS FEATURES
// Full JavaScript Implementation - Ultimate Version
// ==========================================

document.addEventListener('DOMContentLoaded', function() {
    
    // ==========================================
    // GLOBAL VARIABLES
    // ==========================================
    
    // Certificate variables
    let canvas = document.getElementById('sertifikat-canvas');
    let ctx = canvas ? canvas.getContext('2d') : null;
    let certificateImage = null;
    let textElements = [];
    let selectedTextIndex = -1;
    let isDragging = false;
    let dragOffset = { x: 0, y: 0 };
    let scale = 1;
    let history = []; // Undo/Redo history
    let historyIndex = -1;
    let maxHistory = 20;
    
    // Excel Cleaner variables
    let excelData = null;
    let excelFileName = '';
    let selectedNameColumns = [];
    let detectedNameColumns = [];
    
    // NEW: QR Code variables
    let qrCodeElements = [];
    
    // NEW: Image overlay variables
    let imageOverlays = [];
    
    // NEW: Batch rename variables
    let renamePatterns = [];
    
    // NEW: Statistics
    let appStats = {
        certificatesGenerated: 0,
        filesCleaned: 0,
        timeSaved: 0
    };

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
        return 'id_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
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

    function debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }

    // ==========================================
    // HISTORY (UNDO/REDO)
    // ==========================================
    
    function saveHistory() {
        // Remove future history if we're not at the end
        if (historyIndex < history.length - 1) {
            history = history.slice(0, historyIndex + 1);
        }
        
        history.push(JSON.parse(JSON.stringify(textElements)));
        
        if (history.length > maxHistory) {
            history.shift();
        } else {
            historyIndex++;
        }
        
        updateHistoryButtons();
    }

    function undo() {
        if (historyIndex > 0) {
            historyIndex--;
            textElements = JSON.parse(JSON.stringify(history[historyIndex]));
            selectedTextIndex = -1;
            updateToolbar();
            redrawCanvas();
            updateHistoryButtons();
        }
    }

    function redo() {
        if (historyIndex < history.length - 1) {
            historyIndex++;
            textElements = JSON.parse(JSON.stringify(history[historyIndex]));
            selectedTextIndex = -1;
            updateToolbar();
            redrawCanvas();
            updateHistoryButtons();
        }
    }

    function updateHistoryButtons() {
        const undoBtn = document.getElementById('undo-btn');
        const redoBtn = document.getElementById('redo-btn');
        if (undoBtn) undoBtn.disabled = historyIndex <= 0;
        if (redoBtn) redoBtn.disabled = historyIndex >= history.length - 1;
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
                    saveHistory(); // Save initial state
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
                        saveHistory();
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
                dataLink: '',
                type: 'text'
            }, false);
        }
    }

    // ==========================================
    // SECTION 2: CANVAS & TEXT EDITING
    // ==========================================
    
    function redrawCanvas() {
        if (!ctx || !certificateImage) return;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(certificateImage, 0, 0);
        
        // Draw image overlays first (behind text)
        imageOverlays.forEach(img => {
            if (img.visible) {
                ctx.drawImage(img.image, img.x, img.y, img.width, img.height);
            }
        });
        
        // Draw QR codes
        qrCodeElements.forEach(qr => {
            if (qr.visible) {
                ctx.drawImage(qr.image, qr.x, qr.y, qr.size, qr.size);
            }
        });
        
        // Draw text elements
        textElements.forEach((textEl, index) => {
            drawTextElement(textEl, index === selectedTextIndex);
        });
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
        
        // Add shadow if enabled
        if (textEl.shadow) {
            ctx.shadowColor = 'rgba(0,0,0,0.3)';
            ctx.shadowBlur = 4;
            ctx.shadowOffsetX = 2;
            ctx.shadowOffsetY = 2;
        }
        
        const lines = displayText.split('\n');
        const lineHeight = textEl.fontSize * 1.2;
        const totalHeight = lines.length * lineHeight;
        let startY = textEl.y - (totalHeight / 2) + (lineHeight / 2);
        
        lines.forEach((line, i) => {
            ctx.fillText(line, textEl.x, startY + (i * lineHeight));
        });
        
        // Draw selection box
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

    // Canvas mouse events
    if (canvas) {
        canvas.addEventListener('mousedown', handleCanvasMouseDown);
        canvas.addEventListener('mousemove', handleCanvasMouseMove);
        canvas.addEventListener('mouseup', handleCanvasMouseUp);
        canvas.addEventListener('dblclick', handleCanvasDoubleClick);
        
        // Touch support
        canvas.addEventListener('touchstart', handleTouchStart, {passive: false});
        canvas.addEventListener('touchmove', handleTouchMove, {passive: false});
        canvas.addEventListener('touchend', handleTouchEnd);
    }

    function getMousePos(evt) {
        const rect = canvas.getBoundingClientRect();
        const clientX = evt.clientX || (evt.touches && evt.touches[0].clientX);
        const clientY = evt.clientY || (evt.touches && evt.touches[0].clientY);
        return {
            x: (clientX - rect.left) / scale,
            y: (clientY - rect.top) / scale
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
        if (isDragging) {
            saveHistory();
        }
        isDragging = false;
    }

    function handleCanvasDoubleClick(e) {
        const pos = getMousePos(e);
        
        for (let i = textElements.length - 1; i >= 0; i--) {
            if (isPointInText(pos.x, pos.y, textElements[i])) {
                const newText = prompt('Edit text:', textElements[i].text);
                if (newText !== null) {
                    textElements[i].text = newText;
                    saveHistory();
                    redrawCanvas();
                    updateToolbarValues();
                }
                return;
            }
        }
    }

    // Touch handlers
    function handleTouchStart(e) {
        e.preventDefault();
        handleCanvasMouseDown(e);
    }

    function handleTouchMove(e) {
        e.preventDefault();
        handleCanvasMouseMove(e);
    }

    function handleTouchEnd(e) {
        e.preventDefault();
        handleCanvasMouseUp();
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
    const undoBtn = document.getElementById('undo-btn');
    const redoBtn = document.getElementById('redo-btn');

    if (undoBtn) undoBtn.addEventListener('click', undo);
    if (redoBtn) redoBtn.addEventListener('click', redo);

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
                dataLink: '',
                shadow: false,
                type: 'text'
            });
        });
    }

    if (deleteTextBtn) {
        deleteTextBtn.addEventListener('click', function() {
            if (selectedTextIndex !== -1) {
                textElements.splice(selectedTextIndex, 1);
                selectedTextIndex = -1;
                saveHistory();
                updateToolbar();
                redrawCanvas();
            }
        });
    }

    function addTextElement(config, saveHist = true) {
        config.id = generateId();
        textElements.push(config);
        selectedTextIndex = textElements.length - 1;
        updateToolbar();
        redrawCanvas();
        if (saveHist) saveHistory();
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
        if (textShadow) textShadow.checked = el.shadow || false;
        
        if (fontBold) fontBold.classList.toggle('active', el.bold);
        if (fontItalic) fontItalic.classList.toggle('active', el.italic);
    }

    // Text inputs
    const textInput = document.getElementById('text-input');
    const fontFamily = document.getElementById('font-family');
    const fontSize = document.getElementById('font-size');
    const fontColor = document.getElementById('font-color');
    const fontColorHex = document.getElementById('font-color-hex');
    const fontBold = document.getElementById('font-bold');
    const fontItalic = document.getElementById('font-italic');
    const fontAlign = document.getElementById('font-align');
    const textTransform = document.getElementById('text-transform');
    const textShadow = document.getElementById('text-shadow');
    const dataLinkSelect = document.getElementById('data-link-select');

    if (textInput) {
        textInput.addEventListener('input', debounce(function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].text = this.value;
                redrawCanvas();
            }
        }, 300));
    }

    if (fontFamily) {
        fontFamily.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].fontFamily = this.value;
                saveHistory();
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
                saveHistory();
                redrawCanvas();
            }
        });
    }

    if (fontBold) {
        fontBold.addEventListener('click', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].bold = !textElements[selectedTextIndex].bold;
                this.classList.toggle('active');
                saveHistory();
                redrawCanvas();
            }
        });
    }

    if (fontItalic) {
        fontItalic.addEventListener('click', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].italic = !textElements[selectedTextIndex].italic;
                this.classList.toggle('active');
                saveHistory();
                redrawCanvas();
            }
        });
    }

    if (fontAlign) {
        fontAlign.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].align = this.value;
                saveHistory();
                redrawCanvas();
            }
        });
    }

    if (textTransform) {
        textTransform.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].transform = this.value;
                saveHistory();
                redrawCanvas();
            }
        });
    }

    if (textShadow) {
        textShadow.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].shadow = this.checked;
                saveHistory();
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

    if (moveUp) moveUp.addEventListener('click', () => { moveSelectedText(0, -5); saveHistory(); });
    if (moveDown) moveDown.addEventListener('click', () => { moveSelectedText(0, 5); saveHistory(); });
    if (moveLeft) moveLeft.addEventListener('click', () => { moveSelectedText(-5, 0); saveHistory(); });
    if (moveRight) moveRight.addEventListener('click', () => { moveSelectedText(5, 0); saveHistory(); });

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
                    showStatus('bulk-status', '✅ Font berhasil ditambahkan!', 'success');
                }).catch(err => {
                    showStatus('bulk-status', '❌ Error load font: ' + err.message, 'error');
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
                qrCodeElements: qrCodeElements,
                imageOverlays: imageOverlays,
                canvasWidth: canvas ? canvas.width : 0,
                canvasHeight: canvas ? canvas.height : 0,
                version: '2.0',
                savedAt: new Date().toISOString()
            };
            
            const blob = new Blob([JSON.stringify(template, null, 2)], {type: 'application/json'});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'certificate-template-v2.json';
            a.click();
            URL.revokeObjectURL(url);
            
            showStatus('bulk-status', '✅ Template tersimpan!', 'success');
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
                    qrCodeElements = template.qrCodeElements || [];
                    imageOverlays = template.imageOverlays || [];
                    selectedTextIndex = -1;
                    saveHistory();
                    updateToolbar();
                    redrawCanvas();
                    showStatus('bulk-status', '✅ Template berhasil dimuat!', 'success');
                } catch (err) {
                    showStatus('bulk-status', '❌ Error: ' + err.message, 'error');
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
    const downloadJpgBtn = document.getElementById('download-jpg-btn');

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

    if (downloadJpgBtn) {
        downloadJpgBtn.addEventListener('click', function() {
            if (!certificateImage) {
                alert('Upload desain sertifikat terlebih dahulu!');
                return;
            }
            
            const link = document.createElement('a');
            link.download = 'sertifikat.jpg';
            link.href = canvas.toDataURL('image/jpeg', 0.9);
            link.click();
        });
    }

    // ==========================================
    // SECTION 7: BULK GENERATION
    // ==========================================
    
    const bulkFileUpload = document.getElementById('bulk-file-upload');
    const gdocLink = document.getElementById('gdoc-link');
    const fetchGdocBtn = document.getElementById('fetch-gdoc-btn');
    const generateBulkBtn = document.getElementById('generate-bulk-btn');
    const bulkFormat = document.getElementById('bulk-format');
    const downloadAsZip = document.getElementById('download-as-zip');
    const bulkStatus = document.getElementById('bulk-status');

    let bulkData = [];
    let bulkHeaders = [];

    if (bulkFileUpload) {
        bulkFileUpload.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(event) {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
                    
                    processBulkData(jsonData);
                    showStatus('bulk-status', `✅ Data dimuat! ${jsonData.length - 1} baris ditemukan.`, 'success');
                } catch (err) {
                    showStatus('bulk-status', '❌ Error: ' + err.message, 'error');
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
                showStatus('bulk-status', `✅ Data diambil! ${jsonData.length - 1} baris ditemukan.`, 'success');
            } catch (err) {
                showStatus('bulk-status', '❌ Error: ' + err.message, 'error');
            }
        });
    }

    function processBulkData(data) {
        if (data.length < 2) {
            showStatus('bulk-status', 'Data terlalu sedikit', 'error');
            return;
        }
        
        bulkHeaders = data[0];
        bulkData = data.slice(1);
        
        if (dataLinkSelect) {
            dataLinkSelect.innerHTML = '<option value="">-- Tidak dihubungkan --</option>';
            bulkHeaders.forEach((header) => {
                const option = document.createElement('option');
                option.value = header;
                option.textContent = header;
                dataLinkSelect.appendChild(option);
            });
        }
        
        if (generateBulkBtn) generateBulkBtn.disabled = false;
    }

    if (dataLinkSelect) {
        dataLinkSelect.addEventListener('change', function() {
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].dataLink = this.value;
                saveHistory();
            }
        });
    }

    if (generateBulkBtn) {
        generateBulkBtn.addEventListener('click', async function() {
            if (!certificateImage || bulkData.length === 0) {
                showStatus('bulk-status', 'Upload desain dan data terlebih dahulu!', 'error');
                return;
            }
            
            const format = bulkFormat ? bulkFormat.value : 'png';
            const asZip = downloadAsZip ? downloadAsZip.checked : true;
            
            generateBulkBtn.disabled = true;
            const originalText = generateBulkBtn.textContent;
            generateBulkBtn.innerHTML = '<span class="loading-spinner"></span> Memproses...';
            
            const startTime = Date.now();
            
            try {
                if (asZip && format === 'png') {
                    await generateBulkZip();
                } else if (format === 'pdf') {
                    await generateBulkPDF(asZip);
                } else {
                    await generateBulkPNG();
                }
                
                const timeTaken = ((Date.now() - startTime) / 1000).toFixed(1);
                appStats.certificatesGenerated += bulkData.length;
                appStats.timeSaved += bulkData.length * 0.5; // Assume 30 sec manual each
                
                showStatus('bulk-status', `✅ ${bulkData.length} sertifikat selesai dalam ${timeTaken}s!`, 'success', 8000);
                updateStats();
            } catch (err) {
                showStatus('bulk-status', '❌ Error: ' + err.message, 'error');
            } finally {
                generateBulkBtn.disabled = false;
                generateBulkBtn.textContent = originalText;
            }
        });
    }

    async function generateBulkZip() {
        const zip = new JSZip();
        const folder = zip.folder("sertifikat");
        
        for (let i = 0; i < bulkData.length; i++) {
            generateCertificateWithData(bulkData[i], i);
            const dataUrl = canvas.toDataURL('image/png');
            const base64Data = dataUrl.split(',')[1];
            folder.file(`sertifikat_${i + 1}.png`, base64Data, {base64: true});
            
            if (i % 10 === 0) {
                showStatus('bulk-status', `Memproses... ${i + 1}/${bulkData.length}`, 'info', 0);
                await new Promise(r => setTimeout(r, 10));
            }
        }
        
        const content = await zip.generateAsync({type: "blob"});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(content);
        link.download = 'sertifikat.zip';
        link.click();
    }

    async function generateBulkPNG() {
        for (let i = 0; i < bulkData.length; i++) {
            generateCertificateWithData(bulkData[i], i);
            const link = document.createElement('a');
            link.download = `sertifikat_${i + 1}.png`;
            link.href = canvas.toDataURL('image/png');
            link.click();
            
            await new Promise(r => setTimeout(r, 100));
        }
    }

    async function generateBulkPDF(asZip) {
        const { jsPDF } = window.jspdf;
        
        if (asZip) {
            const zip = new JSZip();
            
            for (let i = 0; i < bulkData.length; i++) {
                generateCertificateWithData(bulkData[i], i);
                const pdf = createPDFfromCanvas();
                const pdfBlob = pdf.output('blob');
                zip.file(`sertifikat_${i + 1}.pdf`, pdfBlob);
                
                if (i % 10 === 0) {
                    showStatus('bulk-status', `Memproses... ${i + 1}/${bulkData.length}`, 'info', 0);
                    await new Promise(r => setTimeout(r, 10));
                }
            }
            
            const content = await zip.generateAsync({type: "blob"});
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = 'sertifikat-pdf.zip';
            link.click();
        } else {
            const pdf = new jsPDF({
                orientation: canvas.width > canvas.height ? 'l' : 'p',
                unit: 'px',
                format: [canvas.width, canvas.height]
            });
            
            for (let i = 0; i < bulkData.length; i++) {
                generateCertificateWithData(bulkData[i], i);
                const imgData = canvas.toDataURL('image/png');
                
                if (i > 0) pdf.addPage();
                pdf.addImage(imgData, 'PNG', 0, 0, canvas.width, canvas.height);
                
                if (i % 10 === 0) {
                    showStatus('bulk-status', `Memproses... ${i + 1}/${bulkData.length}`, 'info', 0);
                    await new Promise(r => setTimeout(r, 10));
                }
            }
            
            pdf.save('sertifikat-bulk.pdf');
        }
    }

    function generateCertificateWithData(rowData, index) {
        if (!ctx || !certificateImage) return;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(certificateImage, 0, 0);
        
        // Draw image overlays
        imageOverlays.forEach(img => {
            if (img.visible) {
                ctx.drawImage(img.image, img.x, img.y, img.width, img.height);
            }
        });
        
        // Draw QR codes with data
        qrCodeElements.forEach(qr => {
            if (qr.visible) {
                // Update QR content if linked to data
                if (qr.dataLink) {
                    const colIndex = bulkHeaders.indexOf(qr.dataLink);
                    if (colIndex !== -1 && rowData[colIndex]) {
                        generateQRCode(qr, String(rowData[colIndex]));
                    }
                }
                ctx.drawImage(qr.image, qr.x, qr.y, qr.size, qr.size);
            }
        });
        
        // Draw text elements
        textElements.forEach(textEl => {
            let displayText = textEl.text;
            
            if (textEl.dataLink) {
                const colIndex = bulkHeaders.indexOf(textEl.dataLink);
                if (colIndex !== -1 && rowData[colIndex]) {
                    displayText = String(rowData[colIndex]);
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
            
            if (textEl.shadow) {
                ctx.shadowColor = 'rgba(0,0,0,0.3)';
                ctx.shadowBlur = 4;
                ctx.shadowOffsetX = 2;
                ctx.shadowOffsetY = 2;
            }
            
            const lines = displayText.split('\n');
            const lineHeight = textEl.fontSize * 1.2;
            const totalHeight = lines.length * lineHeight;
            let startY = textEl.y - (totalHeight / 2) + (lineHeight / 2);
            
            lines.forEach((line, i) => {
                ctx.fillText(line, textEl.x, startY + (i * lineHeight));
            });
            
            ctx.restore();
        });
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
    // SECTION 8: SMART EXCEL CLEANER
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
            showStatus('cleaner-status', '❌ Format tidak didukung', 'error');
            return;
        }

        excelFileName = file.name.replace(ext, '');
        if (outputFilename) outputFilename.value = excelFileName + '_clean';
        
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
                
                if (jsonData.length === 0) {
                    showStatus('cleaner-status', '❌ File kosong', 'error');
                    return;
                }

                excelData = jsonData;
                processCleanerData(jsonData);
                
                if (cleanerDropZone) cleanerDropZone.classList.add('has-file');
                if (cleanerWorkspace) cleanerWorkspace.classList.remove('hidden');
                showStatus('cleaner-status', `✅ File dimuat! ${detectedNameColumns.length} kolom nama terdeteksi.`, 'success');
                
            } catch (err) {
                showStatus('cleaner-status', '❌ Error: ' + err.message, 'error');
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
        }
    }

    function populateColumnSelect(headers) {
        if (!nameColumnSelect) return;
        
        nameColumnSelect.innerHTML = '';
        
        const selectAllDiv = document.createElement('div');
        selectAllDiv.className = 'column-option select-all';
        selectAllDiv.innerHTML = `
            <label class="checkbox-label">
                <input type="checkbox" id="select-all-columns">
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
                    <input type="checkbox" value="${index}" ${isSelected ? 'checked' : ''}>
                    <span class="checkmark"></span>
                    <span class="column-name">${header || `Kolom ${index + 1}`}</span>
                    ${isDetected ? '<span class="detected-badge">🎯 Nama</span>' : ''}
                </label>
            `;
            nameColumnSelect.appendChild(div);
        });
        
        const selectAllCheckbox = document.getElementById('select-all-columns');
        if (selectAllCheckbox) {
            selectAllCheckbox.addEventListener('change', function() {
                const checkboxes = nameColumnSelect.querySelectorAll('input[type="checkbox"]:not(#select-all-columns)');
                checkboxes.forEach(cb => cb.checked = this.checked);
                updateSelectedColumnsFromCheckboxes();
            });
        }
        
        const checkboxes = nameColumnSelect.querySelectorAll('input[type="checkbox"]:not(#select-all-columns)');
        checkboxes.forEach(cb => {
            cb.addEventListener('change', updateSelectedColumnsFromCheckboxes);
        });
    }

    function updateSelectedColumnsFromCheckboxes() {
        selectedNameColumns = [];
        const checkboxes = nameColumnSelect.querySelectorAll('input[type="checkbox"]:not(#select-all-columns)');
        checkboxes.forEach(cb => {
            if (cb.checked) selectedNameColumns.push(parseInt(cb.value));
        });
        
        updateCleanerPreview();
        updateSelectedColumnsInfo();
        updateCleanerStats(excelData[0].length);
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
            if (cleanerPreviewBody) cleanerPreviewBody.innerHTML = '<tr><td colspan="2">Pilih kolom</td></tr>';
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
                tr.innerHTML = `<td>${i}</td><td class="empty-cell">-</td>`;
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
            tr.innerHTML = `<td colspan="2" style="text-align: center; color: #6c757d;">... dan ${excelData.length - 11} baris lainnya</td>`;
            cleanerPreviewBody.appendChild(tr);
        }
    }

    if (autoDetectBtn) {
        autoDetectBtn.addEventListener('click', function() {
            if (!excelData) return;
            
            const headers = excelData[0];
            selectedNameColumns = [];
            
            headers.forEach((header, index) => {
                if (isNameColumn(header)) selectedNameColumns.push(index);
            });
            
            updateColumnSelection();
            updateCleanerPreview();
            updateSelectedColumnsInfo();
            updateCleanerStats(headers.length);
            
            showStatus('cleaner-status', `✅ ${selectedNameColumns.length} kolom terdeteksi`, 'success');
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
        });
    }

    if (downloadCleanBtn) {
        downloadCleanBtn.addEventListener('click', function() {
            if (!excelData || selectedNameColumns.length === 0) {
                showStatus('cleaner-status', '❌ Pilih minimal 1 kolom', 'error');
                return;
            }

            const format = outputFormat ? outputFormat.value : 'xlsx';
            const filename = outputFilename ? (outputFilename.value || 'nama_bersih') : 'nama_bersih';
            
            const cleanData = [['Nama']];
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
                appStats.filesCleaned++;
                updateStats();
                showStatus('cleaner-status', `✅ ${totalNames} nama diekspor!`, 'success');
            } catch (err) {
                showStatus('cleaner-status', '❌ Error: ' + err.message, 'error');
            }
        });
    }

    // ==========================================
    // SECTION 9: QR CODE GENERATOR
    // ==========================================
    
    const addQrBtn = document.getElementById('add-qr-btn');
    const qrContentInput = document.getElementById('qr-content');
    const qrSizeInput = document.getElementById('qr-size');
    const qrDataLink = document.getElementById('qr-data-link');

    function generateQRCode(qrObj, content) {
        // Simple QR generation using canvas (basic implementation)
        // For production, use qrcode.js library
        const qrCanvas = document.createElement('canvas');
        qrCanvas.width = qrObj.size;
        qrCanvas.height = qrObj.size;
        const qrCtx = qrCanvas.getContext('2d');
        
        // Draw placeholder QR pattern
        qrCtx.fillStyle = 'white';
        qrCtx.fillRect(0, 0, qrObj.size, qrObj.size);
        qrCtx.fillStyle = 'black';
        
        const cellSize = qrObj.size / 25;
        for (let i = 0; i < 25; i++) {
            for (let j = 0; j < 25; j++) {
                if (Math.random() > 0.5) {
                    qrCtx.fillRect(i * cellSize, j * cellSize, cellSize, cellSize);
                }
            }
        }
        
        // Draw content indicator
        qrCtx.fillStyle = '#1a237e';
        qrCtx.font = '10px Arial';
        qrCtx.textAlign = 'center';
        qrCtx.fillText('QR', qrObj.size/2, qrObj.size/2);
        
        qrObj.image = qrCanvas;
        qrObj.content = content;
    }

    if (addQrBtn) {
        addQrBtn.addEventListener('click', function() {
            if (!canvas) return;
            
            const size = parseInt(qrSizeInput ? qrSizeInput.value : 100);
            const content = qrContentInput ? qrContentInput.value : 'https://example.com';
            const dataLink = qrDataLink ? qrDataLink.value : '';
            
            const qrObj = {
                id: generateId(),
                x: canvas.width / 2,
                y: canvas.height / 2,
                size: size,
                content: content,
                dataLink: dataLink,
                visible: true,
                image: null
            };
            
            generateQRCode(qrObj, content);
            qrCodeElements.push(qrObj);
            saveHistory();
            redrawCanvas();
            
            showStatus('bulk-status', '✅ QR Code ditambahkan!', 'success');
        });
    }

    // ==========================================
    // SECTION 10: IMAGE OVERLAY
    // ==========================================
    
    const addImageBtn = document.getElementById('add-image-btn');
    const overlayImageUpload = document.getElementById('overlay-image-upload');

    if (addImageBtn && overlayImageUpload) {
        addImageBtn.addEventListener('click', () => overlayImageUpload.click());
        
        overlayImageUpload.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(event) {
                const img = new Image();
                img.onload = function() {
                    const overlay = {
                        id: generateId(),
                        image: img,
                        x: canvas.width / 2 - 50,
                        y: canvas.height / 2 - 50,
                        width: 100,
                        height: 100 * (img.height / img.width),
                        visible: true
                    };
                    imageOverlays.push(overlay);
                    saveHistory();
                    redrawCanvas();
                    showStatus('bulk-status', '✅ Gambar ditambahkan!', 'success');
                };
                img.src = event.target.result;
            };
            reader.readAsDataURL(file);
            this.value = '';
        });
    }

    // ==========================================
    // SECTION 11: BATCH RENAME TOOL
    // ==========================================
    
    const renamePatternInput = document.getElementById('rename-pattern');
    const previewRenameBtn = document.getElementById('preview-rename-btn');
    const applyRenameBtn = document.getElementById('apply-rename-btn');
    const renamePreview = document.getElementById('rename-preview');

    if (previewRenameBtn) {
        previewRenameBtn.addEventListener('click', function() {
            if (!bulkData.length) {
                showStatus('bulk-status', 'Load data terlebih dahulu!', 'error');
                return;
            }
            
            const pattern = renamePatternInput ? renamePatternInput.value : 'Sertifikat_{nama}';
            const preview = [];
            
            for (let i = 0; i < Math.min(bulkData.length, 5); i++) {
                let filename = pattern;
                bulkHeaders.forEach((header, idx) => {
                    const regex = new RegExp(`{${header}}`, 'g');
                    filename = filename.replace(regex, String(bulkData[i][idx] || ''));
                });
                filename = filename.replace(/{index}/g, i + 1);
                preview.push(filename);
            }
            
            if (renamePreview) {
                renamePreview.innerHTML = '<strong>Preview:</strong><br>' + 
                    preview.map((p, i) => `${i + 1}. ${p}.png`).join('<br>');
            }
        });
    }

    // ==========================================
    // SECTION 12: KEYBOARD SHORTCUTS
    // ==========================================
    
    document.addEventListener('keydown', function(e) {
        // Ctrl/Cmd + Z = Undo
        if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
            e.preventDefault();
            undo();
        }
        
        // Ctrl/Cmd + Shift + Z = Redo
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'z') {
            e.preventDefault();
            redo();
        }
        
        // Ctrl/Cmd + S = Save Template
        if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            if (saveTemplateBtn) saveTemplateBtn.click();
        }
        
        // Delete = Delete selected text
        if (e.key === 'Delete' && selectedTextIndex !== -1) {
            if (deleteTextBtn) deleteTextBtn.click();
        }
        
        // Arrow keys = Move selected text
        if (selectedTextIndex !== -1) {
            const step = e.shiftKey ? 10 : 1;
            switch(e.key) {
                case 'ArrowUp': 
                    e.preventDefault();
                    moveSelectedText(0, -step); 
                    break;
                case 'ArrowDown': 
                    e.preventDefault();
                    moveSelectedText(0, step); 
                    break;
                case 'ArrowLeft': 
                    e.preventDefault();
                    moveSelectedText(-step, 0); 
                    break;
                case 'ArrowRight': 
                    e.preventDefault();
                    moveSelectedText(step, 0); 
                    break;
            }
        }
    });

    // ==========================================
    // SECTION 13: STATISTICS DASHBOARD
    // ==========================================
    
    function updateStats() {
        const statsCerts = document.getElementById('stats-certificates');
        const statsFiles = document.getElementById('stats-files');
        const statsTime = document.getElementById('stats-time');
        
        if (statsCerts) statsCerts.textContent = appStats.certificatesGenerated.toLocaleString();
        if (statsFiles) statsFiles.textContent = appStats.filesCleaned.toLocaleString();
        if (statsTime) statsTime.textContent = Math.floor(appStats.timeSaved / 60) + ' jam';
    }

    // ==========================================
    // SECTION 14: AUTO SAVE
    // ==========================================
    
    setInterval(function() {
        if (textElements.length > 0) {
            localStorage.setItem('certificate_autosave', JSON.stringify({
                textElements: textElements,
                timestamp: new Date().toISOString()
            }));
        }
    }, 30000); // Auto save every 30 seconds

    // Restore auto-saved data on load
    const autosaved = localStorage.getItem('certificate_autosave');
    if (autosaved) {
        try {
            const data = JSON.parse(autosaved);
            const age = Date.now() - new Date(data.timestamp).getTime();
            if (age < 24 * 60 * 60 * 1000) { // Less than 24 hours
                if (confirm('Ada data tersimpan otomatis. Pulihkan?')) {
                    textElements = data.textElements || [];
                    updateToolbar();
                    redrawCanvas();
                }
            }
        } catch(e) {}
    }

    // ==========================================
    // SECTION 15: COLOR PICKER FROM CANVAS
    // ==========================================
    
    const colorPickerBtn = document.getElementById('color-picker-btn');
    let isPickingColor = false;

    if (colorPickerBtn) {
        colorPickerBtn.addEventListener('click', function() {
            isPickingColor = !isPickingColor;
            this.classList.toggle('active', isPickingColor);
            if (canvas) canvas.style.cursor = isPickingColor ? 'crosshair' : 'default';
            
            if (isPickingColor) {
                showStatus('bulk-status', 'Klik pada gambar untuk mengambil warna', 'info');
            }
        });
    }

    if (canvas) {
        canvas.addEventListener('click', function(e) {
            if (!isPickingColor || !certificateImage) return;
            
            const pos = getMousePos(e);
            const pixel = ctx.getImageData(pos.x, pos.y, 1, 1).data;
            const hex = '#' + ('000000' + rgbToHex(pixel[0], pixel[1], pixel[2])).slice(-6);
            
            if (selectedTextIndex !== -1) {
                textElements[selectedTextIndex].color = hex;
                if (fontColor) fontColor.value = hex;
                if (fontColorHex) fontColorHex.value = hex;
                saveHistory();
                redrawCanvas();
            }
            
            isPickingColor = false;
            if (colorPickerBtn) colorPickerBtn.classList.remove('active');
            canvas.style.cursor = 'default';
        });
    }

    function rgbToHex(r, g, b) {
        return ((r << 16) | (g << 8) | b).toString(16);
    }

    // ==========================================
    // INITIALIZATION
    // ==========================================
    
    console.log('🚀 Certificate Generator Ultimate v2.0 initialized!');
    console.log('💡 Tips: Gunakan Ctrl+Z untuk undo, arrow keys untuk geser teks');
    
    updateHistoryButtons();
    updateStats();
    
}); // End DOMContentLoaded
