// ==========================================
// CERTIFICATE GENERATOR PRO + SMART EXCEL CLEANER
// Full JavaScript Implementation with 5 New Features
// Version: 2.0 Pro
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
    
    // Excel Cleaner variables
    let excelData = null;
    let excelFileName = '';
    let selectedNameColumns = [];
    let detectedNameColumns = [];
    
    // NEW: Feature 1 - Image Library
    let imageLibrary = JSON.parse(localStorage.getItem('certImageLibrary') || '[]');
    
    // NEW: Feature 2 - QR Code settings
    let qrCodeSettings = {
        enabled: false,
        dataPattern: '{{name}}-{{id}}',
        position: 'bottom-right',
        size: 100
    };
    
    // NEW: Feature 3 - Analytics
    let analytics = JSON.parse(localStorage.getItem('certAnalytics') || '{"totalGenerated":0,"history":[]}');
    
    // NEW: Feature 4 - Batch Style
    let batchStyle = {
        target: 'all', // 'all', 'selected', 'unlinked', 'linked'
        properties: []
    };
    
    // NEW: Feature 5 - Preview Data
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
                    
                    // NEW: Auto-save to library option
                    showSaveToLibraryOption(file.name, event.target.result);
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
                        
                        // Convert PDF to image for library
                        showSaveToLibraryOption(file.name.replace('.pdf', '.png'), tempCanvas.toDataURL());
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
        
        // NEW: Initialize new features
        initImageLibrary();
        initQRCodeFeature();
        initAnalytics();
        initBatchStyleEditor();
        initPreviewMode();
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
    // NEW FEATURE 1: IMAGE LIBRARY 📚
    // ==========================================
    
    function showSaveToLibraryOption(filename, dataUrl) {
        // Create temporary notification
        const notif = document.createElement('div');
        notif.className = 'library-notification';
        notif.innerHTML = `
            <div class="notif-content">
                <span>💾 Simpan ke Library?</span>
                <button id="save-lib-yes" class="btn-primary btn-small">Ya</button>
                <button id="save-lib-no" class="btn-secondary btn-small">Tidak</button>
            </div>
        `;
        document.body.appendChild(notif);
        
        document.getElementById('save-lib-yes').addEventListener('click', () => {
            saveToLibrary(filename, dataUrl);
            notif.remove();
        });
        
        document.getElementById('save-lib-no').addEventListener('click', () => {
            notif.remove();
        });
        
        setTimeout(() => notif.remove(), 10000);
    }

    function saveToLibrary(name, dataUrl) {
        const id = 'img_' + Date.now();
        const thumb = createThumbnail(dataUrl, 200);
        
        imageLibrary.push({
            id: id,
            name: name,
            thumbnail: thumb,
            fullImage: dataUrl,
            dateAdded: new Date().toISOString(),
            useCount: 0
        });
        
        localStorage.setItem('certImageLibrary', JSON.stringify(imageLibrary));
        showStatus('bulk-status', '✅ Desain tersimpan di Library!', 'success');
        updateLibraryUI();
    }

    function createThumbnail(dataUrl, maxSize) {
        // Return same image for now (in production, resize it)
        return dataUrl;
    }

    function initImageLibrary() {
        const libContainer = document.getElementById('image-library-container');
        if (!libContainer) return;
        
        updateLibraryUI();
        
        // Setup event listeners for library buttons
        document.getElementById('open-library-btn')?.addEventListener('click', openLibraryModal);
        document.getElementById('clear-library-btn')?.addEventListener('click', clearLibrary);
    }

    function updateLibraryUI() {
        const grid = document.getElementById('library-grid');
        if (!grid) return;
        
        grid.innerHTML = '';
        
        if (imageLibrary.length === 0) {
            grid.innerHTML = '<div class="empty-library">Library kosong. Upload desain untuk menyimpan.</div>';
            return;
        }
        
        // Sort by most used and recent
        const sorted = [...imageLibrary].sort((a, b) => b.useCount - a.useCount || new Date(b.dateAdded) - new Date(a.dateAdded));
        
        sorted.forEach(item => {
            const div = document.createElement('div');
            div.className = 'library-item';
            div.innerHTML = `
                <img src="${item.thumbnail}" alt="${item.name}" loading="lazy">
                <div class="item-overlay">
                    <span class="use-count">Used: ${item.useCount}x</span>
                    <button class="use-item-btn" data-id="${item.id}">Gunakan</button>
                    <button class="delete-item-btn" data-id="${item.id}">🗑️</button>
                </div>
                <div class="item-name">${item.name}</div>
            `;
            grid.appendChild(div);
        });
        
        // Add event listeners
        grid.querySelectorAll('.use-item-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                useLibraryImage(btn.dataset.id);
            });
        });
        
        grid.querySelectorAll('.delete-item-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                deleteLibraryImage(btn.dataset.id);
            });
        });
    }

    function useLibraryImage(id) {
        const item = imageLibrary.find(img => img.id === id);
        if (!item) return;
        
        const img = new Image();
        img.onload = function() {
            certificateImage = img;
            initCanvas();
            redrawCanvas();
            
            // Update use count
            item.useCount++;
            localStorage.setItem('certImageLibrary', JSON.stringify(imageLibrary));
            updateLibraryUI();
            
            showStatus('bulk-status', `✅ Desain "${item.name}" dimuat!`, 'success');
            closeLibraryModal();
        };
        img.src = item.fullImage;
    }

    function deleteLibraryImage(id) {
        if (!confirm('Hapus desain dari library?')) return;
        
        imageLibrary = imageLibrary.filter(img => img.id !== id);
        localStorage.setItem('certImageLibrary', JSON.stringify(imageLibrary));
        updateLibraryUI();
        showStatus('bulk-status', 'Desain dihapus dari library', 'info');
    }

    function openLibraryModal() {
        const modal = document.getElementById('library-modal');
        if (modal) modal.classList.remove('hidden');
        updateLibraryUI();
    }

    function closeLibraryModal() {
        const modal = document.getElementById('library-modal');
        if (modal) modal.classList.add('hidden');
    }

    function clearLibrary() {
        if (!confirm('Hapus SEMUA desain di library? Ini tidak bisa dibatalkan.')) return;
        imageLibrary = [];
        localStorage.removeItem('certImageLibrary');
        updateLibraryUI();
        showStatus('bulk-status', 'Library dikosongkan', 'info');
    }

    // ==========================================
    // NEW FEATURE 2: QR CODE GENERATOR 📱
    // ==========================================
    
    function initQRCodeFeature() {
        const qrToggle = document.getElementById('qr-enable');
        const qrPattern = document.getElementById('qr-pattern');
        const qrPosition = document.getElementById('qr-position');
        const qrSize = document.getElementById('qr-size');
        const previewBtn = document.getElementById('qr-preview-btn');
        
        if (qrToggle) {
            qrToggle.addEventListener('change', function() {
                qrCodeSettings.enabled = this.checked;
                updateQRPreview();
            });
        }
        
        if (qrPattern) {
            qrPattern.addEventListener('input', function() {
                qrCodeSettings.dataPattern = this.value;
            });
        }
        
        if (qrPosition) {
            qrPosition.addEventListener('change', function() {
                qrCodeSettings.position = this.value;
                updateQRPreview();
            });
        }
        
        if (qrSize) {
            qrSize.addEventListener('input', function() {
                qrCodeSettings.size = parseInt(this.value);
                updateQRPreview();
            });
        }
        
        if (previewBtn) {
            previewBtn.addEventListener('click', generateQRPreview);
        }
    }

    function updateQRPreview() {
        if (!qrCodeSettings.enabled || !certificateImage) return;
        redrawCanvas(); // Will include QR if enabled
    }

    function generateQRCode(data, size) {
        // Simple QR-like pattern (in production, use qrcode.js library)
        const qrCanvas = document.createElement('canvas');
        qrCanvas.width = size;
        qrCanvas.height = size;
        const qrCtx = qrCanvas.getContext('2d');
        
        // Generate pseudo-random pattern based on data
        const seed = data.split('').reduce((a,b)=>a+b.charCodeAt(0),0);
        
        qrCtx.fillStyle = 'white';
        qrCtx.fillRect(0, 0, size, size);
        qrCtx.fillStyle = 'black';
        
        const cellSize = Math.floor(size / 25);
        
        // Position detection patterns (corners)
        drawPositionPattern(qrCtx, 0, 0, cellSize * 7);
        drawPositionPattern(qrCtx, size - cellSize * 7, 0, cellSize * 7);
        drawPositionPattern(qrCtx, 0, size - cellSize * 7, cellSize * 7);
        
        // Data pattern
        for (let i = 0; i < 25; i++) {
            for (let j = 0; j < 25; j++) {
                // Skip position detection areas
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

    function generateQRPreview() {
        const sampleData = 'Sample-Name-001';
        const qr = generateQRCode(sampleData, qrCodeSettings.size);
        
        const modal = document.getElementById('qr-preview-modal');
        const container = document.getElementById('qr-preview-container');
        
        if (container) {
            container.innerHTML = '';
            container.appendChild(qr);
        }
        
        if (modal) modal.classList.remove('hidden');
    }

    // ==========================================
    // NEW FEATURE 3: ANALYTICS DASHBOARD 📊
    // ==========================================
    
    function initAnalytics() {
        updateAnalyticsUI();
        
        document.getElementById('view-analytics-btn')?.addEventListener('click', openAnalyticsModal);
        document.getElementById('export-analytics-btn')?.addEventListener('click', exportAnalytics);
        document.getElementById('clear-analytics-btn')?.addEventListener('click', clearAnalytics);
    }

    function recordGeneration(count, format) {
        analytics.totalGenerated += count;
        analytics.history.push({
            date: new Date().toISOString(),
            count: count,
            format: format,
            template: textElements.length > 0 ? 'Custom' : 'Default'
        });
        
        // Keep only last 100 records
        if (analytics.history.length > 100) {
            analytics.history = analytics.history.slice(-100);
        }
        
        localStorage.setItem('certAnalytics', JSON.stringify(analytics));
        updateAnalyticsUI();
    }

    function updateAnalyticsUI() {
        const totalEl = document.getElementById('analytics-total');
        const todayEl = document.getElementById('analytics-today');
        const weekEl = document.getElementById('analytics-week');
        const chartEl = document.getElementById('analytics-chart');
        
        if (totalEl) totalEl.textContent = analytics.totalGenerated.toLocaleString();
        
        // Calculate today's count
        const today = new Date().toDateString();
        const todayCount = analytics.history.filter(h => new Date(h.date).toDateString() === today)
            .reduce((sum, h) => sum + h.count, 0);
        if (todayEl) todayEl.textContent = todayCount.toLocaleString();
        
        // Calculate this week's count
        const weekAgo = new Date();
        weekAgo.setDate(weekAgo.getDate() - 7);
        const weekCount = analytics.history.filter(h => new Date(h.date) > weekAgo)
            .reduce((sum, h) => sum + h.count, 0);
        if (weekEl) weekEl.textContent = weekCount.toLocaleString();
        
        // Simple bar chart
        if (chartEl) {
            drawSimpleChart(chartEl, analytics.history);
        }
    }

    function drawSimpleChart(container, history) {
        // Group by date (last 7 days)
        const days = {};
        for (let i = 6; i >= 0; i--) {
            const d = new Date();
            d.setDate(d.getDate() - i);
            days[d.toLocaleDateString('id-ID', {weekday:'short'})] = 0;
        }
        
        history.forEach(h => {
            const date = new Date(h.date);
            const dayKey = date.toLocaleDateString('id-ID', {weekday:'short'});
            if (days.hasOwnProperty(dayKey)) {
                days[dayKey] += h.count;
            }
        });
        
        const max = Math.max(...Object.values(days), 1);
        const bars = Object.entries(days).map(([day, count]) => {
            const height = (count / max) * 100;
            return `
                <div class="chart-bar-wrapper">
                    <div class="chart-bar" style="height: ${height}%">
                        <span class="bar-value">${count}</span>
                    </div>
                    <span class="bar-label">${day}</span>
                </div>
            `;
        }).join('');
        
        container.innerHTML = `<div class="simple-chart">${bars}</div>`;
    }

    function openAnalyticsModal() {
        const modal = document.getElementById('analytics-modal');
        if (modal) {
            updateAnalyticsUI();
            modal.classList.remove('hidden');
        }
    }

    function exportAnalytics() {
        const data = {
            exportDate: new Date().toISOString(),
            analytics: analytics,
            summary: {
                total: analytics.totalGenerated,
                records: analytics.history.length
            }
        };
        
        const blob = new Blob([JSON.stringify(data, null, 2)], {type: 'application/json'});
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `analytics-${new Date().toISOString().split('T')[0]}.json`;
        a.click();
        URL.revokeObjectURL(url);
    }

    function clearAnalytics() {
        if (!confirm('Hapus semua data analytics?')) return;
        analytics = { totalGenerated: 0, history: [] };
        localStorage.removeItem('certAnalytics');
        updateAnalyticsUI();
    }

    // ==========================================
    // NEW FEATURE 4: BATCH STYLE EDITOR 🎨
    // ==========================================
    
    function initBatchStyleEditor() {
        const targetSelect = document.getElementById('batch-target');
        const propertySelect = document.getElementById('batch-property');
        const valueInput = document.getElementById('batch-value');
        const applyBtn = document.getElementById('batch-apply-btn');
        const previewBtn = document.getElementById('batch-preview-btn');
        
        if (applyBtn) {
            applyBtn.addEventListener('click', applyBatchStyle);
        }
        
        if (previewBtn) {
            previewBtn.addEventListener('click', previewBatchStyle);
        }
        
        // Populate property select
        if (propertySelect) {
            const properties = [
                {value: 'fontSize', label: 'Ukuran Font'},
                {value: 'fontFamily', label: 'Font Family'},
                {value: 'color', label: 'Warna'},
                {value: 'bold', label: 'Bold (true/false)'},
                {value: 'italic', label: 'Italic (true/false)'},
                {value: 'align', label: 'Alignment (left/center/right)'},
                {value: 'transform', label: 'Transform (none/uppercase/titlecase)'}
            ];
            
            propertySelect.innerHTML = properties.map(p => 
                `<option value="${p.value}">${p.label}</option>`
            ).join('');
        }
    }

    function applyBatchStyle() {
        const target = document.getElementById('batch-target')?.value || 'all';
        const property = document.getElementById('batch-property')?.value;
        let value = document.getElementById('batch-value')?.value;
        
        if (!property || value === '') {
            showStatus('bulk-status', 'Pilih property dan masukkan nilai', 'error');
            return;
        }
        
        // Convert value types
        if (property === 'fontSize') value = parseInt(value);
        if (property === 'bold' || property === 'italic') value = value === 'true' || value === '1';
        
        let affected = 0;
        
        textElements.forEach((el, index) => {
            let shouldApply = false;
            
            switch(target) {
                case 'all': shouldApply = true; break;
                case 'selected': shouldApply = index === selectedTextIndex; break;
                case 'unlinked': shouldApply = !el.dataLink; break;
                case 'linked': shouldApply = !!el.dataLink; break;
            }
            
            if (shouldApply) {
                el[property] = value;
                affected++;
            }
        });
        
        redrawCanvas();
        updateToolbarValues();
        showStatus('bulk-status', `✅ ${affected} elemen diupdate!`, 'success');
    }

    function previewBatchStyle() {
        // Create temporary preview without applying
        const property = document.getElementById('batch-property')?.value;
        let value = document.getElementById('batch-value')?.value;
        
        if (!property || !value) return;
        
        // Store original values
        const originals = textElements.map(el => ({...el}));
        
        // Apply temporarily
        applyBatchStyle();
        
        // Show modal with before/after
        setTimeout(() => {
            // Restore after 3 seconds
            textElements = originals;
            redrawCanvas();
            showStatus('bulk-status', 'Preview selesai (style dikembalikan)', 'info');
        }, 3000);
    }

    // ==========================================
    // NEW FEATURE 5: PREVIEW MODE 🔍
    // ==========================================
    
    function initPreviewMode() {
        const openPreviewBtn = document.getElementById('open-preview-btn');
        const closePreviewBtn = document.getElementById('close-preview-modal');
        const prevBtn = document.getElementById('preview-prev');
        const nextBtn = document.getElementById('preview-next');
        const useDataBtn = document.getElementById('use-preview-data');
        
        if (openPreviewBtn) {
            openPreviewBtn.addEventListener('click', openPreviewModal);
        }
        
        if (closePreviewBtn) {
            closePreviewBtn.addEventListener('click', closePreviewModal);
        }
        
        if (prevBtn) {
            prevBtn.addEventListener('click', () => navigatePreview(-1));
        }
        
        if (nextBtn) {
            nextBtn.addEventListener('click', () => navigatePreview(1));
        }
        
        if (useDataBtn) {
            useDataBtn.addEventListener('click', usePreviewDataForBulk);
        }
    }

    function openPreviewModal() {
        if (!certificateImage) {
            showStatus('bulk-status', 'Upload desain terlebih dahulu!', 'error');
            return;
        }
        
        if (!bulkData || bulkData.length === 0) {
            showStatus('bulk-status', 'Upload data Excel terlebih dahulu!', 'error');
            return;
        }
        
        previewData = bulkData;
        previewIndex = 0;
        
        updatePreviewUI();
        
        const modal = document.getElementById('preview-modal');
        if (modal) modal.classList.remove('hidden');
    }

    function closePreviewModal() {
        const modal = document.getElementById('preview-modal');
        if (modal) modal.classList.add('hidden');
        
        // Restore original canvas
        redrawCanvas();
    }

    function updatePreviewUI() {
        const counter = document.getElementById('preview-counter');
        const canvas = document.getElementById('preview-canvas');
        
        if (counter) {
            counter.textContent = `${previewIndex + 1} / ${previewData.length}`;
        }
        
        // Generate certificate with current preview data
        generateCertificateWithData(previewData[previewIndex]);
        
        // Copy to preview canvas
        if (canvas) {
            const pCtx = canvas.getContext('2d');
            canvas.width = canvas.width;
            canvas.height = canvas.height;
            pCtx.drawImage(canvas, 0, 0);
        }
    }

    function navigatePreview(direction) {
        previewIndex += direction;
        
        if (previewIndex < 0) previewIndex = previewData.length - 1;
        if (previewIndex >= previewData.length) previewIndex = 0;
        
        updatePreviewUI();
    }

    function usePreviewDataForBulk() {
        // Confirm this data for bulk generation
        showStatus('bulk-status', `✅ Data preview siap! ${previewData.length} sertifikat akan dibuat.`, 'success');
        closePreviewModal();
        
        // Scroll to bulk section
        document.getElementById('bulk-section')?.scrollIntoView({behavior: 'smooth'});
    }

    // ==========================================
    // SECTION 2: CANVAS & TEXT EDITING (Updated with QR)
    // ==========================================
    
    function redrawCanvas() {
        if (!ctx || !certificateImage) return;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(certificateImage, 0, 0);
        
        // Draw text elements
        textElements.forEach((textEl, index) => {
            drawTextElement(textEl, index === selectedTextIndex);
        });
        
        // NEW: Draw QR Code if enabled
        if (qrCodeSettings.enabled) {
            drawQRCodeOnCanvas();
        }
    }

    function drawQRCodeOnCanvas() {
        // Generate sample QR or use preview data
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

    // Event listeners for text editing
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
                    
                    // Restore QR settings if present
                    if (template.qrSettings) {
                        qrCodeSettings = template.qrSettings;
                    }
                    
                    selectedTextIndex = -1;
                    updateToolbar();
                    redrawCanvas();
                    
                    // Update QR UI
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
    // SECTION 7: BULK GENERATION (Updated with Analytics)
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
        bulkData = data.slice(1);
        
        if (dataLinkSelect) {
            dataLinkSelect.innerHTML = '<option value="">-- Tidak dihubungkan --</option>';
            bulkHeaders.forEach((header, index) => {
                const option = document.createElement('option');
                option.value = header;
                option.textContent = header;
                dataLinkSelect.appendChild(option);
            });
        }
        
        if (generateBulkBtn) {
            generateBulkBtn.disabled = false;
        }
        
        // NEW: Update preview data
        previewData = bulkData;
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
            if (!certificateImage || bulkData.length === 0) {
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
                
                // NEW: Record analytics
                recordGeneration(bulkData.length, format);
                
                showStatus('bulk-status', '✅ Semua sertifikat berhasil dibuat!', 'success');
            } catch (err) {
                showStatus('bulk-status', '❌ Error: ' + err.message, 'error');
            } finally {
                generateBulkBtn.disabled = false;
                generateBulkBtn.textContent = 'Mulai Generasi Massal';
            }
        });
    }

    async function generateBulkZip() {
        const zip = new JSZip();
        const folder = zip.folder("sertifikat");
        
        for (let i = 0; i < bulkData.length; i++) {
            generateCertificateWithData(bulkData[i]);
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
            generateCertificateWithData(bulkData[i]);
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
                generateCertificateWithData(bulkData[i]);
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
                generateCertificateWithData(bulkData[i]);
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

    function generateCertificateWithData(rowData) {
        if (!ctx || !certificateImage) return;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(certificateImage, 0, 0);
        
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
            
            const lines = displayText.split('\n');
            const lineHeight = textEl.fontSize * 1.2;
            const totalHeight = lines.length * lineHeight;
            let startY = textEl.y - (totalHeight / 2) + (lineHeight / 2);
            
            lines.forEach((line, i) => {
                ctx.fillText(line, textEl.x, startY + (i * lineHeight));
            });
            
            ctx.restore();
        });
        
        // Draw QR if enabled
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
            showStatus('cleaner-status', '❌ Format file tidak didukung. Gunakan .csv, .xls, atau .xlsx', 'error');
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
        });
    }

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
    // SECTION 9: COLOR PICKER FROM CANVAS
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
    // MODAL CLOSE HANDLERS
    // ==========================================
    
    // Close modals when clicking outside
    window.addEventListener('click', function(e) {
        if (e.target.classList.contains('modal-overlay')) {
            e.target.classList.add('hidden');
        }
    });

    // Close button handlers for all modals
    document.querySelectorAll('.modal-close-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            this.closest('.modal-overlay')?.classList.add('hidden');
        });
    });

    // ==========================================
    // INITIALIZATION
    // ==========================================
    
    console.log('🚀 Certificate Generator Pro v2.0 initialized!');
    console.log('Features: Image Library, QR Code, Analytics, Batch Style, Preview Mode');
    
}); // End DOMContentLoaded
