// ==========================================
// CERTIFICATE GENERATOR PRO - FINAL FIX
// Fixed: Data filtering & Empty row handling
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
    
    // CRITICAL: Data untuk bulk generation
    let cleanedDataForBulk = []; // Hanya berisi [nama] yang valid
    let bulkHeaders = ['Nama']; // Default untuk data cleaned
    let originalFileName = '';
    
    // Excel Cleaner variables
    let excelData = null;
    let excelFileName = '';
    let selectedNameColumns = [];
    let detectedNameColumns = [];

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
        // Hapus whitespace berlebih dan karakter aneh
        return String(name)
            .trim()
            .replace(/\s+/g, ' ')
            .replace(/[^\w\s\-\'\.]/g, '') // Hanya izinkan alphanumeric, spasi, -, ', .
            .trim();
    }

    function isValidName(name) {
        if (!name) return false;
        const cleaned = cleanName(name);
        // Minimal 2 karakter dan bukan hanya angka/simbol
        return cleaned.length >= 2 && /[a-zA-Z]/.test(cleaned);
    }

    function sanitizeFileName(name) {
        if (!name) return 'Unknown';
        return String(name)
            .replace(/[<>:"/\\|?*]/g, '_')
            .replace(/\s+/g, ' ')
            .trim()
            .substring(0, 100);
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
    // SECTION 3: TOOLBAR CONTROLS (Simplified)
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

    // Event listeners (simplified)
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
    // SECTION 4: TEMPLATE SAVE/LOAD
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
                version: '2.0 Final'
            };
            
            const blob = new Blob([JSON.stringify(template, null, 2)], {type: 'application/json'});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'certificate-template.json';
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
                    selectedTextIndex = -1;
                    updateToolbar();
                    redrawCanvas();
                    alert('Template berhasil dimuat!');
                } catch (err) {
                    alert('Error membaca template: ' + err.message);
                }
            };
            reader.readAsText(file);
            this.value = '';
        });
    }

    // ==========================================
    // SECTION 5: DOWNLOAD SINGLE
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
    // SECTION 6: BULK GENERATION (FINAL FIX)
    // ==========================================
    
    const generateBulkBtn = document.getElementById('generate-bulk-btn');
    const bulkFormat = document.getElementById('bulk-format');
    const downloadAsZip = document.getElementById('download-as-zip');

    // CRITICAL FIX: Tombol generate hanya aktif jika ada data valid
    function updateGenerateButtonState() {
        if (generateBulkBtn) {
            const hasData = cleanedDataForBulk && cleanedDataForBulk.length > 0;
            const hasDesign = !!certificateImage;
            generateBulkBtn.disabled = !(hasData && hasDesign);
            
            console.log('Button state:', {hasData, hasDesign, count: cleanedDataForBulk.length});
        }
    }

    if (generateBulkBtn) {
        generateBulkBtn.addEventListener('click', async function() {
            // VALIDASI KETAT
            if (!certificateImage) {
                showStatus('bulk-status', '❌ Upload desain sertifikat terlebih dahulu!', 'error');
                return;
            }
            
            if (!cleanedDataForBulk || cleanedDataForBulk.length === 0) {
                showStatus('bulk-status', '❌ Tidak ada data nama! Gunakan Excel Cleaner terlebih dahulu.', 'error');
                return;
            }
            
            const format = bulkFormat ? bulkFormat.value : 'png';
            const asZip = downloadAsZip ? downloadAsZip.checked : true;
            
            // Log untuk debug
            console.log('Starting bulk generation:', {
                totalData: cleanedDataForBulk.length,
                sampleData: cleanedDataForBulk.slice(0, 3),
                format: format,
                asZip: asZip
            });
            
            generateBulkBtn.disabled = true;
            generateBulkBtn.textContent = 'Memproses...';
            
            try {
                let successCount = 0;
                
                if (asZip && format === 'png') {
                    successCount = await generateBulkZip();
                } else if (format === 'pdf') {
                    successCount = await generateBulkPDF(asZip);
                } else {
                    successCount = await generateBulkPNG();
                }
                
                showStatus('bulk-status', `✅ ${successCount} sertifikat berhasil dibuat!`, 'success', 8000);
            } catch (err) {
                console.error('Bulk generation error:', err);
                showStatus('bulk-status', '❌ Error: ' + err.message, 'error');
            } finally {
                generateBulkBtn.disabled = false;
                generateBulkBtn.textContent = 'Mulai Generasi Massal';
            }
        });
    }

    // FIXED: Generate ZIP dengan validasi ketat
    async function generateBulkZip() {
        const zip = new JSZip();
        const folder = zip.folder("sertifikat");
        
        let processedCount = 0;
        let validCount = 0;
        
        for (let i = 0; i < cleanedDataForBulk.length; i++) {
            const nameData = cleanedDataForBulk[i];
            
            // CRITICAL: Validasi nama
            if (!nameData || !nameData[0]) {
                console.log(`Skipping row ${i}: empty data`);
                continue;
            }
            
            const name = cleanName(nameData[0]);
            if (!isValidName(name)) {
                console.log(`Skipping row ${i}: invalid name "${name}"`);
                continue;
            }
            
            // Generate sertifikat
            generateCertificateWithData([name]);
            
            const dataUrl = canvas.toDataURL('image/png');
            const base64Data = dataUrl.split(',')[1];
            
            const safeName = sanitizeFileName(name);
            const fileName = `S - ${safeName}.png`;
            
            folder.file(fileName, base64Data, {base64: true});
            
            validCount++;
            processedCount++;
            
            // Update progress setiap 5 item
            if (processedCount % 5 === 0) {
                showStatus('bulk-status', `Memproses... ${processedCount}/${cleanedDataForBulk.length} (Valid: ${validCount})`, 'info', 0);
                await new Promise(r => setTimeout(r, 10));
            }
        }
        
        if (validCount === 0) {
            throw new Error('Tidak ada nama valid untuk digenerate!');
        }
        
        const content = await zip.generateAsync({type: "blob"});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(content);
        link.download = `${originalFileName || 'sertifikat'}_${validCount}items.zip`;
        link.click();
        
        return validCount;
    }

    // FIXED: Generate PNG individual
    async function generateBulkPNG() {
        let validCount = 0;
        
        for (let i = 0; i < cleanedDataForBulk.length; i++) {
            const nameData = cleanedDataForBulk[i];
            
            if (!nameData || !nameData[0]) continue;
            
            const name = cleanName(nameData[0]);
            if (!isValidName(name)) continue;
            
            generateCertificateWithData([name]);
            
            const link = document.createElement('a');
            const safeName = sanitizeFileName(name);
            link.download = `S - ${safeName}.png`;
            link.href = canvas.toDataURL('image/png');
            link.click();
            
            validCount++;
            await new Promise(r => setTimeout(r, 150)); // Delay lebih lama untuk browser
        }
        
        return validCount;
    }

    // FIXED: Generate PDF
    async function generateBulkPDF(asZip) {
        const { jsPDF } = window.jspdf;
        
        if (asZip) {
            const zip = new JSZip();
            let validCount = 0;
            
            for (let i = 0; i < cleanedDataForBulk.length; i++) {
                const nameData = cleanedDataForBulk[i];
                
                if (!nameData || !nameData[0]) continue;
                
                const name = cleanName(nameData[0]);
                if (!isValidName(name)) continue;
                
                generateCertificateWithData([name]);
                
                const pdf = createPDFfromCanvas();
                const pdfBlob = pdf.output('blob');
                
                const safeName = sanitizeFileName(name);
                const fileName = `S - ${safeName}.pdf`;
                
                zip.file(fileName, pdfBlob);
                validCount++;
                
                if (validCount % 5 === 0) {
                    await new Promise(r => setTimeout(r, 10));
                }
            }
            
            if (validCount === 0) throw new Error('Tidak ada nama valid!');
            
            const content = await zip.generateAsync({type: "blob"});
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = `${originalFileName || 'sertifikat'}_${validCount}items_pdf.zip`;
            link.click();
            
            return validCount;
        } else {
            // Single PDF
            const pdf = new jsPDF({
                orientation: canvas.width > canvas.height ? 'l' : 'p',
                unit: 'px',
                format: [canvas.width, canvas.height]
            });
            
            let validCount = 0;
            let firstPage = true;
            
            for (let i = 0; i < cleanedDataForBulk.length; i++) {
                const nameData = cleanedDataForBulk[i];
                
                if (!nameData || !nameData[0]) continue;
                
                const name = cleanName(nameData[0]);
                if (!isValidName(name)) continue;
                
                generateCertificateWithData([name]);
                const imgData = canvas.toDataURL('image/png');
                
                if (!firstPage) pdf.addPage();
                pdf.addImage(imgData, 'PNG', 0, 0, canvas.width, canvas.height);
                firstPage = false;
                validCount++;
            }
            
            if (validCount === 0) throw new Error('Tidak ada nama valid!');
            
            pdf.save(`${originalFileName || 'sertifikat'}_${validCount}items.pdf`);
            return validCount;
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

    function generateCertificateWithData(nameArray) {
        if (!ctx || !certificateImage || !nameArray || !nameArray[0]) return;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(certificateImage, 0, 0);
        
        const name = nameArray[0];
        
        textElements.forEach(textEl => {
            let displayText = textEl.text;
            
            // Jika ada data link, coba gunakan (tapi untuk data cleaned, selalu pakai name)
            if (textEl.dataLink && bulkHeaders.includes(textEl.dataLink)) {
                const colIndex = bulkHeaders.indexOf(textEl.dataLink);
                if (colIndex !== -1 && nameArray[colIndex]) {
                    displayText = String(nameArray[colIndex]);
                }
            } else {
                // Default: gunakan nama utama
                displayText = String(name);
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
    }

    // ==========================================
    // SECTION 7: SMART EXCEL CLEANER (FINAL FIX)
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

    // CRITICAL: Tombol untuk transfer data ke bulk generation
    const useForBulkBtn = document.createElement('button');
    useForBulkBtn.id = 'use-for-bulk-btn';
    useForBulkBtn.className = 'btn-primary btn-large';
    useForBulkBtn.innerHTML = '🚀 Gunakan untuk Generate Massal';
    useForBulkBtn.style.marginTop = '15px';
    useForBulkBtn.style.display = 'none';
    useForBulkBtn.style.background = '#ff6b6b';
    
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
            showStatus('cleaner-status', '❌ Format file tidak didukung', 'error');
            return;
        }

        excelFileName = file.name.replace(ext, '');
        originalFileName = excelFileName;
        
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

                // CRITICAL: Filter baris kosong dari awal
                excelData = jsonData.filter(row => row.some(cell => cell && String(cell).trim() !== ''));
                
                if (excelData.length < 2) {
                    showStatus('cleaner-status', '❌ Tidak ada data valid (minimal perlu header + 1 data)', 'error');
                    return;
                }

                console.log('Excel loaded:', {
                    totalRows: excelData.length,
                    headers: excelData[0],
                    sampleRow: excelData[1]
                });

                processCleanerData(excelData);
                
                if (cleanerDropZone) cleanerDropZone.classList.add('has-file');
                if (cleanerWorkspace) cleanerWorkspace.classList.remove('hidden');
                showStatus('cleaner-status', `✅ File loaded! ${excelData.length - 1} baris data.`, 'success');
                
            } catch (err) {
                showStatus('cleaner-status', '❌ Error: ' + err.message, 'error');
                console.error(err);
            }
        };
        
        reader.readAsArrayBuffer(file);
    }

    function processCleanerData(data) {
        const headers = data[0];
        
        selectedNameColumns = [];
        detectedNameColumns = [];
        
        // Auto-detect kolom nama
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
        updateCleanerStats(headers.length);
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
        
        // Select All
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
        
        // Individual columns
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
        
        // Event listeners
        document.getElementById('select-all-columns')?.addEventListener('change', function() {
            const checkboxes = nameColumnSelect.querySelectorAll('input[type="checkbox"]:not(#select-all-columns)');
            checkboxes.forEach(cb => cb.checked = this.checked);
            updateSelectedColumnsFromCheckboxes();
        });
        
        nameColumnSelect.querySelectorAll('input[type="checkbox"]:not(#select-all-columns)').forEach(cb => {
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
        updateUseForBulkButton();
    }

    function updateUseForBulkButton() {
        const hasData = selectedNameColumns.length > 0 && excelData && excelData.length > 1;
        if (useForBulkBtn) {
            useForBulkBtn.style.display = hasData ? 'block' : 'none';
        }
    }

    // CRITICAL FIX: Transfer data ke bulk generation
    if (useForBulkBtn) {
        useForBulkBtn.addEventListener('click', function() {
            if (!excelData || selectedNameColumns.length === 0) {
                showStatus('cleaner-status', '❌ Pilih kolom nama terlebih dahulu!', 'error');
                return;
            }
            
            // CRITICAL: Build cleaned data dengan FILTERING KETAT
            const tempData = [];
            let skippedCount = 0;
            let emptyCount = 0;
            
            for (let i = 1; i < excelData.length; i++) {
                const row = excelData[i];
                const namesFromRow = [];
                
                // Ambil nama dari semua kolom yang dipilih
                selectedNameColumns.forEach(colIndex => {
                    const rawValue = row[colIndex];
                    const name = cleanName(rawValue);
                    
                    if (name && isValidName(name)) {
                        namesFromRow.push(name);
                    } else if (rawValue) {
                        skippedCount++;
                    } else {
                        emptyCount++;
                    }
                });
                
                // Tambahkan setiap nama valid sebagai entry terpisah
                namesFromRow.forEach(name => {
                    tempData.push([name]);
                });
            }
            
            // VALIDASI FINAL
            if (tempData.length === 0) {
                showStatus('cleaner-status', '❌ Tidak ada nama valid ditemukan! Periksa data Anda.', 'error');
                return;
            }
            
            // Set ke variabel global
            cleanedDataForBulk = tempData;
            bulkHeaders = ['Nama'];
            
            console.log('Data prepared for bulk:', {
                totalValid: cleanedDataForBulk.length,
                sample: cleanedDataForBulk.slice(0, 5),
                skippedInvalid: skippedCount,
                skippedEmpty: emptyCount
            });
            
            // Update UI
            updateGenerateButtonState();
            
            // Scroll ke bulk section
            const bulkSection = document.getElementById('bulk-section');
            if (bulkSection) {
                bulkSection.scrollIntoView({behavior: 'smooth'});
                // Highlight the section
                bulkSection.style.animation = 'highlight 2s';
                setTimeout(() => bulkSection.style.animation = '', 2000);
            }
            
            showStatus('bulk-status', 
                `✅ ${cleanedDataForBulk.length} nama valid siap! ` +
                `(Skip: ${skippedCount} invalid, ${emptyCount} kosong)`, 
                'success', 10000);
        });
    }

    function updateCleanerStats(totalCols) {
        if (totalRowsEl) totalRowsEl.textContent = excelData ? (excelData.length - 1) : 0;
        if (removedColsEl) removedColsEl.textContent = totalCols - selectedNameColumns.length;
        if (keptColsEl) keptColsEl.textContent = selectedNameColumns.length;
    }

    function updateSelectedColumnsInfo() {
        if (!selectedColumnsInfo) return;
        
        if (selectedNameColumns.length === 0) {
            selectedColumnsInfo.innerHTML = '<span class="warning">⚠️ Pilih minimal 1 kolom nama</span>';
            return;
        }
        
        const headers = excelData[0];
        const selectedNames = selectedNameColumns.map(idx => headers[idx] || `Kolom ${idx + 1}`);
        
        selectedColumnsInfo.innerHTML = `
            <strong>${selectedNameColumns.length} kolom dipilih:</strong> ${selectedNames.join(', ')}
        `;
    }

    function updateCleanerPreview() {
        if (!cleanerPreviewBody || !excelData || selectedNameColumns.length === 0) {
            if (cleanerPreviewBody) cleanerPreviewBody.innerHTML = '<tr><td colspan="2" style="text-align: center;">Pilih kolom untuk preview</td></tr>';
            return;
        }
        
        cleanerPreviewBody.innerHTML = '';
        const rowsToShow = Math.min(excelData.length - 1, 10);
        let displayIndex = 1;
        
        for (let i = 1; i <= rowsToShow; i++) {
            const row = excelData[i];
            const validNames = [];
            
            selectedNameColumns.forEach(colIndex => {
                const name = cleanName(row[colIndex]);
                if (isValidName(name)) validNames.push(name);
            });
            
            if (validNames.length === 0) {
                const tr = document.createElement('tr');
                tr.style.opacity = '0.5';
                tr.innerHTML = `<td>${i}</td><td style="color: #999;">- (kosong/invalid)</td>`;
                cleanerPreviewBody.appendChild(tr);
            } else {
                validNames.forEach((name, idx) => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${displayIndex++}</td>
                        <td><strong>${name}</strong></td>
                    `;
                    cleanerPreviewBody.appendChild(tr);
                });
            }
        }
        
        if (excelData.length > 11) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td colspan="2" style="text-align: center; color: #6c757d;">
                    ... ${excelData.length - 11} baris lainnya
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
                if (isNameColumn(header)) selectedNameColumns.push(index);
            });
            
            updateColumnSelection();
            updateCleanerPreview();
            updateSelectedColumnsInfo();
            updateCleanerStats(headers.length);
            updateUseForBulkButton();
            
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
            updateUseForBulkButton();
        });
    }

    if (downloadCleanBtn) {
        downloadCleanBtn.addEventListener('click', function() {
            if (!excelData || selectedNameColumns.length === 0) {
                showStatus('cleaner-status', '❌ Pilih kolom terlebih dahulu', 'error');
                return;
            }

            const format = outputFormat ? outputFormat.value : 'xlsx';
            const filename = outputFilename ? (outputFilename.value || 'nama_bersih') : 'nama_bersih';
            
            const cleanData = [['Nama']];
            let validCount = 0;
            
            for (let i = 1; i < excelData.length; i++) {
                const row = excelData[i];
                
                selectedNameColumns.forEach(colIndex => {
                    const name = cleanName(row[colIndex]);
                    if (isValidName(name)) {
                        cleanData.push([name]);
                        validCount++;
                    }
                });
            }
            
            if (validCount === 0) {
                showStatus('cleaner-status', '❌ Tidak ada nama valid untuk diexport!', 'error');
                return;
            }

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(cleanData);
            ws['!cols'] = [{wch: 50}];
            XLSX.utils.book_append_sheet(wb, ws, "Nama Bersih");

            const ext = format === 'csv' ? '.csv' : '.xlsx';
            XLSX.writeFile(wb, filename + ext);
            showStatus('cleaner-status', `✅ ${validCount} nama diexport!`, 'success');
        });
    }

    // ==========================================
    // INITIALIZATION
    // ==========================================
    
    console.log('✅ Certificate Generator FINAL FIX loaded');
    console.log('Features: Strict filtering, Valid name detection, Proper data flow');
    
    // Disable generate button on start
    updateGenerateButtonState();
    
}); // End DOMContentLoaded
