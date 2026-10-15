// ===================================================================
// PDF & IMAGE MASTER — MASTER APPLICATION JAVASCRIPT
// ===================================================================

document.addEventListener('DOMContentLoaded', () => {

    // ==========================================
    // 1. TOOL DEFINITIONS & CONFIGURATION
    // ==========================================
    const TOOLS = {
        'pdf-to-word': {
            id: 'pdf-to-word',
            title: 'PDF เป็น Word (.docx)',
            desc: 'แปลงข้อความและเค้าโครงจาก PDF เป็นเอกสาร Word ที่แก้ไขได้',
            icon: 'fas fa-file-word',
            accept: '.pdf,application/pdf',
            multiple: false,
            endpoint: '/api/pdf/to-word',
            options: []
        },
        'pdf-to-images': {
            id: 'pdf-to-images',
            title: 'PDF เป็น รูปภาพ (PNG/JPG)',
            desc: 'แปลงแต่ละหน้าของ PDF เป็นภาพความละเอียดสูง',
            icon: 'fas fa-file-image',
            accept: '.pdf,application/pdf',
            multiple: false,
            clientSide: true, // Uses HTML5 canvas + pdf.js/pdf-lib
            options: [
                { id: 'imgFormat', label: 'ฟอร์แมตรูปภาพ', type: 'select', options: [
                    { value: 'png', text: 'PNG (คมชัด ไม่สูญเสียคุณภาพ)' },
                    { value: 'jpeg', text: 'JPG (ไฟล์ขนาดกะทัดรัด)' },
                    { value: 'webp', text: 'WebP (ทันสมัย ขนาดเล็ก)' }
                ], default: 'png' },
                { id: 'imgDpi', label: 'ความละเอียด (Scale)', type: 'select', options: [
                    { value: '1.5', text: 'มาตรฐาน (150 DPI)' },
                    { value: '2.0', text: 'คมชัดสูง (200 DPI)' },
                    { value: '3.0', text: 'คมชัดสูงสุด (300 DPI)' }
                ], default: '2.0' }
            ]
        },
        'pdf-to-text': {
            id: 'pdf-to-text',
            title: 'PDF เป็น ข้อความ (.txt)',
            desc: 'สกัดข้อความทั้งหมดออกมาเป็นไฟล์ Text',
            icon: 'fas fa-file-lines',
            accept: '.pdf,application/pdf',
            multiple: false,
            endpoint: '/api/pdf/to-text',
            options: []
        },
        'image-to-pdf': {
            id: 'image-to-pdf',
            title: 'รูปภาพเป็น PDF',
            desc: 'นำรูปภาพหลายรูปมาจัดรวมเป็นเล่ม PDF',
            icon: 'fas fa-images',
            accept: 'image/png,image/jpeg,image/webp,image/gif,image/bmp',
            multiple: true,
            endpoint: '/api/image/to-pdf',
            options: [
                { id: 'pageSize', label: 'ขนาดหน้ากระดาษ', type: 'select', options: [
                    { value: 'a4', text: 'กระดาษมาตรฐาน A4' },
                    { value: 'fit', text: 'พอดีกับขนาดรูปภาพแต่ละรูป (Fit to Image)' }
                ], default: 'a4' },
                { id: 'orientation', label: 'แนววางกระดาษ', type: 'select', options: [
                    { value: 'portrait', text: 'แนวตั้ง (Portrait)' },
                    { value: 'landscape', text: 'แนวนอน (Landscape)' }
                ], default: 'portrait' }
            ]
        },
        'pdf-merge': {
            id: 'pdf-merge',
            title: 'รวมไฟล์ PDF',
            desc: 'รวมหลายไฟล์ PDF ให้เป็นไฟล์เดียว',
            icon: 'fas fa-object-group',
            accept: '.pdf,application/pdf',
            multiple: true,
            endpoint: '/api/pdf/merge',
            options: []
        },
        'pdf-split': {
            id: 'pdf-split',
            title: 'แยกไฟล์ PDF',
            desc: 'แยกเอกสารเป็นไฟล์ละหน้า หรือแยกตามช่วงที่ต้องการ',
            icon: 'fas fa-scissors',
            accept: '.pdf,application/pdf',
            multiple: false,
            endpoint: '/api/pdf/split',
            options: [
                { id: 'splitMode', label: 'รูปแบบการแยก', type: 'select', options: [
                    { value: 'all', text: 'แยกทุกหน้า (1 หน้าต่อ 1 ไฟล์)' },
                    { value: 'range', text: 'กำหนดช่วงหน้า (เช่น 1-3, 5, 7-10)' }
                ], default: 'all' },
                { id: 'splitRange', label: 'ระบุช่วงหน้า (เมื่อเลือกกำหนดช่วง)', type: 'text', placeholder: 'เช่น 1-3, 5, 7-9', default: '' }
            ]
        },
        'pdf-extract': {
            id: 'pdf-extract',
            title: 'ดึงเฉพาะหน้า PDF',
            desc: 'ดึงเฉพาะหน้าที่ต้องการออกมาเป็น PDF เล่มใหม่',
            icon: 'fas fa-file-export',
            accept: '.pdf,application/pdf',
            multiple: false,
            endpoint: '/api/pdf/extract-pages',
            options: [
                { id: 'extractPages', label: 'ระบุหน้าที่ต้องการดึง', type: 'text', placeholder: 'เช่น 1, 3, 5-8', default: '1' }
            ]
        },
        'pdf-delete': {
            id: 'pdf-delete',
            title: 'ลบหน้า PDF ที่ไม่ต้องการ',
            desc: 'ตัดหน้าที่เลือกทิ้งออกจากเอกสารอย่างรวดเร็ว',
            icon: 'fas fa-trash-can',
            accept: '.pdf,application/pdf',
            multiple: false,
            endpoint: '/api/pdf/delete-pages',
            options: [
                { id: 'deletePages', label: 'ระบุหน้าที่ต้องการลบ', type: 'text', placeholder: 'เช่น 2, 4-6', default: '' }
            ]
        },
        'pdf-rotate': {
            id: 'pdf-rotate',
            title: 'หมุนหน้า PDF',
            desc: 'หมุนหน้าเอกสาร 90, 180, 270 องศา',
            icon: 'fas fa-rotate',
            accept: '.pdf,application/pdf',
            multiple: false,
            endpoint: '/api/pdf/rotate',
            options: [
                { id: 'rotateAngle', label: 'องศาการหมุน', type: 'select', options: [
                    { value: '90', text: '90° ตามเข็มนาฬิกา' },
                    { value: '180', text: '180° กลับหัว' },
                    { value: '270', text: '270° ทวนเข็มนาฬิกา' }
                ], default: '90' },
                { id: 'rotatePages', label: 'หน้าที่ต้องการหมุน', type: 'text', placeholder: 'พิมพ์ all หรือระบุเช่น 1, 3-5', default: 'all' }
            ]
        },
        'pdf-compress': {
            id: 'pdf-compress',
            title: 'บีบอัดลดขนาด PDF',
            desc: 'ลดขนาดไฟล์ PDF ให้เล็กลงเพื่อความสะดวกในการส่งต่อ',
            icon: 'fas fa-compress',
            accept: '.pdf,application/pdf',
            multiple: false,
            endpoint: '/api/pdf/compress',
            options: []
        },
        'pdf-watermark': {
            id: 'pdf-watermark',
            title: 'ใส่ลายน้ำ PDF',
            desc: 'ประทับข้อความลายน้ำลงบนทุกหน้าของเอกสาร',
            icon: 'fas fa-stamp',
            accept: '.pdf,application/pdf',
            multiple: false,
            endpoint: '/api/pdf/watermark',
            options: [
                { id: 'watermarkText', label: 'ข้อความลายน้ำ', type: 'text', placeholder: 'เช่น ลับเฉพาะ, DRAFT', default: 'CONFIDENTIAL' },
                { id: 'watermarkOpacity', label: 'ความโปร่งแสง', type: 'range', min: 10, max: 80, step: 5, default: 25, unit: '%' }
            ]
        },
        'pdf-page-numbers': {
            id: 'pdf-page-numbers',
            title: 'ใส่เลขหน้า PDF',
            desc: 'เพิ่มหมายเลขหน้าลงที่ด้านล่างของทุกหน้าโดยอัตโนมัติ',
            icon: 'fas fa-list-ol',
            accept: '.pdf,application/pdf',
            multiple: false,
            endpoint: '/api/pdf/page-numbers',
            options: [
                { id: 'pageNumberFormat', label: 'รูปแบบเลขหน้า', type: 'select', options: [
                    { value: 'Page {n} of {total}', text: 'Page {n} of {total}' },
                    { value: 'หน้า {n} / {total}', text: 'หน้า {n} / {total}' },
                    { value: '- {n} -', text: '- {n} -' }
                ], default: 'Page {n} of {total}' }
            ]
        },
        'image-convert': {
            id: 'image-convert',
            title: 'แปลงฟอร์แมตรูปภาพ',
            desc: 'แปลงรูปภาพเป็น PNG, JPG, WebP, AVIF, TIFF',
            icon: 'fas fa-arrows-rotate',
            accept: 'image/*',
            multiple: false,
            endpoint: '/api/image/convert',
            options: [
                { id: 'targetFormat', label: 'แปลงเป็นฟอร์แมต', type: 'select', options: [
                    { value: 'png', text: 'PNG (Lossless โปร่งแสง)' },
                    { value: 'jpg', text: 'JPG / JPEG (มาตรฐาน)' },
                    { value: 'webp', text: 'WebP (ขนาดเล็ก คุณภาพสูง)' },
                    { value: 'avif', text: 'AVIF (เทคโนโลยีใหม่ ขนาดจิ๋ว)' },
                    { value: 'tiff', text: 'TIFF (งานพิมพ์ความละเอียดสูง)' }
                ], default: 'png' }
            ]
        },
        'image-resize': {
            id: 'image-resize',
            title: 'ปรับขนาดรูปภาพ',
            desc: 'ปรับความกว้างและความสูงตามต้องการ',
            icon: 'fas fa-up-right-and-down-left-from-center',
            accept: 'image/*',
            multiple: false,
            endpoint: '/api/image/resize',
            options: [
                { id: 'resizeWidth', label: 'ความกว้าง (Width px)', type: 'number', placeholder: 'เช่น 1920', default: '' },
                { id: 'resizeHeight', label: 'ความสูง (Height px)', type: 'number', placeholder: 'เช่น 1080', default: '' },
                { id: 'maintainRatio', label: 'คงอัตราส่วนภาพเดิมไว้ (Maintain Ratio)', type: 'checkbox', default: true }
            ]
        },
        'image-compress': {
            id: 'image-compress',
            title: 'บีบอัดลดขนาดรูปภาพ',
            desc: 'ลดขนาดรูปภาพโดยยังคงรายละเอียดภาพไว้',
            icon: 'fas fa-file-zipper',
            accept: 'image/*',
            multiple: false,
            endpoint: '/api/image/compress',
            options: [
                { id: 'compressQuality', label: 'ระดับคุณภาพ', type: 'range', min: 10, max: 95, step: 5, default: 70, unit: '%' },
                { id: 'compressFormat', label: 'ฟอร์แมตผลลัพธ์', type: 'select', options: [
                    { value: 'jpeg', text: 'JPEG / JPG' },
                    { value: 'webp', text: 'WebP (แนะนำ)' },
                    { value: 'png', text: 'PNG' }
                ], default: 'jpeg' }
            ]
        },
        'image-rotate': {
            id: 'image-rotate',
            title: 'หมุน & กลับด้านรูปภาพ',
            desc: 'หมุนองศา หรือกลับภาพแนวนอน/แนวตั้ง',
            icon: 'fas fa-arrows-spin',
            accept: 'image/*',
            multiple: false,
            endpoint: '/api/image/rotate',
            options: [
                { id: 'imgAngle', label: 'องศาการหมุน', type: 'select', options: [
                    { value: '0', text: 'ไม่หมุน' },
                    { value: '90', text: 'หมุนตามเข็ม 90°' },
                    { value: '180', text: 'กลับหัว 180°' },
                    { value: '270', text: 'หมุนทวนเข็ม 270°' }
                ], default: '90' },
                { id: 'imgFlipH', label: 'พลิกภาพแนวนอน (Flip Horizontal / กระจกเงา)', type: 'checkbox', default: false },
                { id: 'imgFlipV', label: 'พลิกภาพแนวตั้ง (Flip Vertical)', type: 'checkbox', default: false }
            ]
        }
    };

    // ==========================================
    // 2. STATE VARIABLES
    // ==========================================
    let currentTool = null;
    let selectedFiles = [];

    // DOM Elements
    const themeToggle = document.getElementById('themeToggle');
    const toolSearch = document.getElementById('toolSearch');
    const categoryTabs = document.querySelectorAll('.category-tabs .tab-btn');
    const toolCards = document.querySelectorAll('.tool-card');

    const modalBackdrop = document.getElementById('modalBackdrop');
    const modalCloseBtn = document.getElementById('modalCloseBtn');
    const modalCancelBtn = document.getElementById('modalCancelBtn');
    const modalTitle = document.getElementById('modalTitle');
    const modalDesc = document.getElementById('modalDesc');
    const modalIcon = document.getElementById('modalIcon');

    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const dropzoneHint = document.getElementById('dropzoneHint');

    const filesPreviewContainer = document.getElementById('filesPreviewContainer');
    const filesCountBadge = document.getElementById('filesCountBadge');
    const filesList = document.getElementById('filesList');
    const clearFilesBtn = document.getElementById('clearFilesBtn');

    const optionsPanel = document.getElementById('optionsPanel');
    const optionsGrid = document.getElementById('optionsGrid');

    const processStatusArea = document.getElementById('processStatusArea');
    const progressBarFill = document.getElementById('progressBarFill');
    const statusText = document.getElementById('statusText');
    const statusPercent = document.getElementById('statusPercent');

    const resultCard = document.getElementById('resultCard');
    const downloadBtn = document.getElementById('downloadBtn');
    const fileSizeDiff = document.getElementById('fileSizeDiff');
    const origSizeText = document.getElementById('origSizeText');
    const newSizeText = document.getElementById('newSizeText');
    const savedBadge = document.getElementById('savedBadge');

    const processBtn = document.getElementById('processBtn');
    const toastContainer = document.getElementById('toastContainer');

    // ==========================================
    // 3. THEME TOGGLE
    // ==========================================
    const savedTheme = localStorage.getItem('master_theme') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);

    themeToggle.addEventListener('click', () => {
        const curTheme = document.documentElement.getAttribute('data-theme');
        const nextTheme = curTheme === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', nextTheme);
        localStorage.setItem('master_theme', nextTheme);
        showToast('info', `เปลี่ยนเป็นโหมด ${nextTheme === 'dark' ? 'มืด' : 'สว่าง'} เรียบร้อย`);
    });

    // ==========================================
    // 4. CATEGORY & SEARCH FILTER
    // ==========================================
    categoryTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            categoryTabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            filterTools();
        });
    });

    toolSearch.addEventListener('input', () => {
        filterTools();
    });

    // Auto-focus search on load for fast typing
    if (window.innerWidth > 600 && toolSearch) {
        setTimeout(() => toolSearch.focus(), 250);
    }

    // Keyboard Shortcuts: Ctrl+K or / to jump to search bar immediately
    document.addEventListener('keydown', (e) => {
        const isModalOpen = modalBackdrop && modalBackdrop.classList.contains('active');
        if (isModalOpen) return;

        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
            e.preventDefault();
            toolSearch.focus();
            toolSearch.select();
        } else if (e.key === '/' && document.activeElement !== toolSearch && !['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) {
            e.preventDefault();
            toolSearch.focus();
            toolSearch.select();
        } else if (e.key === 'Escape' && document.activeElement === toolSearch) {
            toolSearch.value = '';
            filterTools();
            toolSearch.blur();
        }
    });

    function filterTools() {
        const activeCategory = document.querySelector('.category-tabs .tab-btn.active').dataset.category;
        const query = toolSearch.value.trim().toLowerCase();

        toolCards.forEach(card => {
            const cardCat = card.dataset.category || '';
            const cardTitle = card.querySelector('.tool-name').textContent.toLowerCase();
            const cardDesc = card.querySelector('.tool-desc').textContent.toLowerCase();

            const matchCat = activeCategory === 'all' || cardCat.includes(activeCategory);
            const matchSearch = !query || cardTitle.includes(query) || cardDesc.includes(query);

            card.style.display = (matchCat && matchSearch) ? 'flex' : 'none';
        });
    }

    // ==========================================
    // 5. MODAL OPEN & CLOSE
    // ==========================================
    toolCards.forEach(card => {
        card.addEventListener('click', () => {
            const toolId = card.dataset.tool;
            openModal(toolId);
        });
    });

    function openModal(toolId) {
        const tool = TOOLS[toolId];
        if (!tool) return;

        currentTool = tool;
        selectedFiles = [];

        // Set Headers
        modalTitle.textContent = tool.title;
        modalDesc.textContent = tool.desc;
        modalIcon.innerHTML = `<i class="${tool.icon}"></i>`;
        dropzoneHint.textContent = tool.multiple
            ? 'รองรับการเลือกหลายไฟล์พร้อมกัน (ขนาดรวมสูงสุด 100MB)'
            : 'เลือกหรือลากไฟล์มาวาง 1 ไฟล์ (ขนาดสูงสุด 100MB)';

        fileInput.accept = tool.accept;
        fileInput.multiple = tool.multiple;
        fileInput.value = '';

        // Render Options
        renderOptions(tool.options);

        // Reset Status & Preview
        resetModalState();

        // Show Modal
        modalBackdrop.classList.add('active');
        document.body.style.overflow = 'hidden';
    }

    function closeModal() {
        modalBackdrop.classList.remove('active');
        document.body.style.overflow = '';
        currentTool = null;
        selectedFiles = [];
        resetModalState();
    }

    modalCloseBtn.addEventListener('click', closeModal);
    modalCancelBtn.addEventListener('click', closeModal);
    modalBackdrop.addEventListener('click', (e) => {
        if (e.target === modalBackdrop) closeModal();
    });

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && modalBackdrop.classList.contains('active')) {
            closeModal();
        }
    });

    function resetModalState() {
        filesPreviewContainer.style.display = 'none';
        filesList.innerHTML = '';
        processStatusArea.style.display = 'none';
        progressBarFill.style.width = '0%';
        resultCard.style.display = 'none';
        processBtn.disabled = true;
        processBtn.style.display = 'inline-flex';
        dropzone.style.display = 'block';
    }

    // ==========================================
    // 6. RENDER DYNAMIC OPTIONS
    // ==========================================
    function renderOptions(options) {
        if (!options || options.length === 0) {
            optionsPanel.style.display = 'none';
            optionsGrid.innerHTML = '';
            return;
        }

        optionsPanel.style.display = 'block';
        optionsGrid.innerHTML = '';

        options.forEach(opt => {
            const group = document.createElement('div');
            group.className = 'form-group';

            if (opt.type === 'select') {
                group.innerHTML = `
                    <label class="form-label" for="${opt.id}">${opt.label}</label>
                    <select class="form-select" id="${opt.id}">
                        ${opt.options.map(o => `<option value="${o.value}" ${o.value === opt.default ? 'selected' : ''}>${o.text}</option>`).join('')}
                    </select>
                `;
            } else if (opt.type === 'range') {
                group.innerHTML = `
                    <label class="form-label" for="${opt.id}">
                        <span>${opt.label}</span>
                        <span class="badge-val" id="${opt.id}Val">${opt.default}${opt.unit || ''}</span>
                    </label>
                    <input type="range" class="form-range" id="${opt.id}" min="${opt.min}" max="${opt.max}" step="${opt.step || 1}" value="${opt.default}">
                `;
            } else if (opt.type === 'checkbox') {
                group.innerHTML = `
                    <label class="checkbox-group" for="${opt.id}">
                        <input type="checkbox" id="${opt.id}" ${opt.default ? 'checked' : ''}>
                        <span>${opt.label}</span>
                    </label>
                `;
            } else {
                group.innerHTML = `
                    <label class="form-label" for="${opt.id}">${opt.label}</label>
                    <input type="${opt.type || 'text'}" class="form-input" id="${opt.id}" placeholder="${opt.placeholder || ''}" value="${opt.default || ''}">
                `;
            }

            optionsGrid.appendChild(group);

            // Real-time badge update for range sliders
            if (opt.type === 'range') {
                const slider = group.querySelector(`#${opt.id}`);
                const badge = group.querySelector(`#${opt.id}Val`);
                slider.addEventListener('input', () => {
                    badge.textContent = `${slider.value}${opt.unit || ''}`;
                });
            }
        });
    }

    // ==========================================
    // 7. FILE UPLOAD & PREVIEW
    // ==========================================
    dropzone.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', (e) => {
        handleFiles(Array.from(e.target.files));
    });

    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleFiles(Array.from(e.dataTransfer.files));
        }
    });

    clearFilesBtn.addEventListener('click', () => {
        selectedFiles = [];
        updateFilesUI();
    });

    function handleFiles(files) {
        if (!currentTool) return;

        if (currentTool.multiple) {
            selectedFiles = [...selectedFiles, ...files];
        } else {
            selectedFiles = files.slice(0, 1);
        }

        updateFilesUI();
    }

    function updateFilesUI() {
        if (selectedFiles.length === 0) {
            filesPreviewContainer.style.display = 'none';
            filesList.innerHTML = '';
            processBtn.disabled = true;
            return;
        }

        filesPreviewContainer.style.display = 'block';
        filesCountBadge.textContent = `ไฟล์ที่เลือก (${selectedFiles.length})`;
        filesList.innerHTML = '';

        selectedFiles.forEach((file, index) => {
            const item = document.createElement('div');
            item.className = 'file-item';

            const isImg = file.type.startsWith('image/');
            const iconHtml = isImg
                ? `<img src="${URL.createObjectURL(file)}" class="file-item-thumb" alt="Preview">`
                : `<i class="fas fa-file-pdf file-item-icon"></i>`;

            item.innerHTML = `
                <div class="file-item-left">
                    ${iconHtml}
                    <div class="file-item-details">
                        <span class="file-item-name" title="${file.name}">${file.name}</span>
                        <span class="file-item-size">${formatBytes(file.size)}</span>
                    </div>
                </div>
                <button class="file-item-remove" data-index="${index}" title="ลบไฟล์นี้">
                    <i class="fas fa-times"></i>
                </button>
            `;
            filesList.appendChild(item);
        });

        // Remove single file
        filesList.querySelectorAll('.file-item-remove').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const idx = parseInt(btn.dataset.index);
                selectedFiles.splice(idx, 1);
                updateFilesUI();
            });
        });

        processBtn.disabled = false;
        resultCard.style.display = 'none';
    }

    // ==========================================
    // 8. PROCESS & DOWNLOAD HANDLER
    // ==========================================
    processBtn.addEventListener('click', async () => {
        if (!currentTool || selectedFiles.length === 0) return;

        processBtn.disabled = true;
        processStatusArea.style.display = 'flex';
        resultCard.style.display = 'none';
        progressBarFill.style.width = '15%';
        statusPercent.textContent = '15%';
        statusText.textContent = 'กำลังเตรียมไฟล์...';

        try {
            // Check if special client-side tool (PDF to Images)
            if (currentTool.id === 'pdf-to-images') {
                await processPdfToImagesClient();
                return;
            }

            // Standard Server-side API processing
            const formData = new FormData();
            const fileFieldName = currentTool.multiple ? 'files' : 'file';

            if (currentTool.multiple) {
                selectedFiles.forEach(f => formData.append(fileFieldName, f));
            } else {
                formData.append(fileFieldName, selectedFiles[0]);
            }

            // Append Options
            if (currentTool.options) {
                currentTool.options.forEach(opt => {
                    const el = document.getElementById(opt.id);
                    if (el) {
                        if (opt.type === 'checkbox') {
                            formData.append(opt.id, el.checked);
                        } else {
                            formData.append(opt.id, el.value);
                        }
                    }
                });
            }

            // Animate progress simulation
            let prog = 25;
            const progressTimer = setInterval(() => {
                if (prog < 88) {
                    prog += Math.floor(Math.random() * 12) + 5;
                    progressBarFill.style.width = `${Math.min(prog, 88)}%`;
                    statusPercent.textContent = `${Math.min(prog, 88)}%`;
                }
            }, 300);

            statusText.textContent = 'กำลังประมวลผลบนเซิร์ฟเวอร์...';

            const response = await fetch(currentTool.endpoint, {
                method: 'POST',
                body: formData
            });

            clearInterval(progressTimer);

            if (!response.ok) {
                let errMsg = 'การประมวลผลล้มเหลว';
                try {
                    const errJson = await response.json();
                    errMsg = errJson.error || errMsg;
                } catch {}
                throw new Error(errMsg);
            }

            progressBarFill.style.width = '100%';
            statusPercent.textContent = '100%';
            statusText.textContent = 'ประมวลผลเสร็จสมบูรณ์!';

            // Get download filename
            let downloadFilename = 'output_file';
            const dispHeader = response.headers.get('Content-Disposition');
            if (dispHeader) {
                const utfMatch = dispHeader.match(/filename\*=UTF-8''([^;]+)/i);
                if (utfMatch) {
                    downloadFilename = decodeURIComponent(utfMatch[1]);
                } else {
                    const regMatch = dispHeader.match(/filename=["']?([^"';]+)["']?/i);
                    if (regMatch) downloadFilename = regMatch[1];
                }
            }

            // Handle size difference headers (e.g. compression)
            const origSize = response.headers.get('X-Original-Size');
            const compSize = response.headers.get('X-Compressed-Size');

            const blob = await response.blob();
            const blobUrl = URL.createObjectURL(blob);

            showSuccessResult(blobUrl, downloadFilename, origSize, compSize);
            showToast('success', `${currentTool.title} สำเร็จแล้ว!`);

        } catch (err) {
            console.error('Process Error:', err);
            processStatusArea.style.display = 'none';
            processBtn.disabled = false;
            showToast('error', err.message || 'เกิดข้อผิดพลาดในการประมวลผล');
        }
    });

    // ==========================================
    // 9. CLIENT-SIDE PDF TO IMAGES (CANVAS + JSZip)
    // ==========================================
    async function processPdfToImagesClient() {
        const file = selectedFiles[0];
        const format = document.getElementById('imgFormat').value || 'png';
        const scale = parseFloat(document.getElementById('imgDpi').value) || 2.0;

        statusText.textContent = 'กำลังโหลดและอ่านเอกสาร PDF...';
        progressBarFill.style.width = '20%';

        // Dynamically load PDF.js if not already present
        if (!window.pdfjsLib) {
            statusText.textContent = 'กำลังเชื่อมต่อเอนจิน PDF...';
            await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }

        const arrayBuffer = await file.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const totalPages = pdf.numPages;

        const zip = new JSZip();
        const baseName = file.name.replace(/\.[^/.]+$/, '');

        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
            statusText.textContent = `กำลังเรนเดอร์ภาพหน้าที่ ${pageNum} จาก ${totalPages}...`;
            const percent = Math.floor(20 + (pageNum / totalPages) * 70);
            progressBarFill.style.width = `${percent}%`;
            statusPercent.textContent = `${percent}%`;

            const page = await pdf.getPage(pageNum);
            const viewport = page.getViewport({ scale: scale });

            const canvas = document.createElement('canvas');
            const context = canvas.getContext('2d');
            canvas.width = viewport.width;
            canvas.height = viewport.height;

            await page.render({ canvasContext: context, viewport: viewport }).promise;

            const mime = format === 'png' ? 'image/png' : (format === 'webp' ? 'image/webp' : 'image/jpeg');
            const imgBlob = await new Promise(resolve => canvas.toBlob(resolve, mime, 0.92));
            const imgBuffer = await imgBlob.arrayBuffer();

            zip.file(`${baseName}_page_${pageNum}.${format}`, imgBuffer);
        }

        statusText.textContent = 'กำลังบีบอัดแพ็กเกจไฟล์ ZIP...';
        progressBarFill.style.width = '95%';
        statusPercent.textContent = '95%';

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const zipUrl = URL.createObjectURL(zipBlob);

        progressBarFill.style.width = '100%';
        statusPercent.textContent = '100%';
        statusText.textContent = 'แปลงรูปภาพทุกหน้าเรียบร้อย!';

        showSuccessResult(zipUrl, `${baseName}_images.zip`);
        showToast('success', `แปลง PDF ทั้งหมด ${totalPages} หน้าเรียบร้อย!`);
    }

    // ==========================================
    // 10. SUCCESS DISPLAY
    // ==========================================
    function showSuccessResult(url, filename, origSizeBytes, newSizeBytes) {
        processStatusArea.style.display = 'none';
        resultCard.style.display = 'flex';
        downloadBtn.href = url;
        downloadBtn.download = filename;

        if (origSizeBytes && newSizeBytes) {
            fileSizeDiff.style.display = 'flex';
            origSizeText.textContent = `เดิม: ${formatBytes(parseInt(origSizeBytes))}`;
            newSizeText.textContent = `ใหม่: ${formatBytes(parseInt(newSizeBytes))}`;

            const orig = parseInt(origSizeBytes);
            const comp = parseInt(newSizeBytes);
            if (orig > comp) {
                const savedPct = Math.round(((orig - comp) / orig) * 100);
                savedBadge.textContent = `ประหยัด ${savedPct}%`;
                savedBadge.style.display = 'inline-block';
            } else {
                savedBadge.style.display = 'none';
            }
        } else {
            fileSizeDiff.style.display = 'none';
        }

        processBtn.disabled = false;
    }

    // ==========================================
    // 11. HELPER FUNCTIONS
    // ==========================================
    function formatBytes(bytes) {
        if (!bytes || bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }

    function showToast(type, message) {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;

        let icon = 'fas fa-info-circle';
        if (type === 'success') icon = 'fas fa-circle-check';
        if (type === 'error') icon = 'fas fa-circle-exclamation';

        toast.innerHTML = `
            <i class="${icon} toast-icon"></i>
            <span class="toast-message">${message}</span>
        `;

        toastContainer.appendChild(toast);

        setTimeout(() => {
            toast.classList.add('hiding');
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    }

    function loadScript(src) {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = src;
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

});
