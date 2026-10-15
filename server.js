const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const sharp = require('sharp');
const { PDFDocument, rgb, degrees, StandardFonts } = require('pdf-lib');
const docx = require('docx');
const pdfParse = require('pdf-parse');
const JSZip = require('jszip');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

// Static file serving (check both public directory and root)
const publicDir = path.join(__dirname, 'public');
if (fs.existsSync(publicDir)) {
    app.use(express.static(publicDir));
}
app.use(express.static(__dirname));

// Explicit route for homepage
app.get('/', (req, res) => {
    const publicIndex = path.join(__dirname, 'public', 'index.html');
    const rootIndex = path.join(__dirname, 'index.html');
    
    if (fs.existsSync(publicIndex)) {
        return res.sendFile(publicIndex);
    } else if (fs.existsSync(rootIndex)) {
        return res.sendFile(rootIndex);
    } else {
        let dirFiles = [];
        try {
            dirFiles = fs.readdirSync(__dirname);
        } catch (e) {
            dirFiles = [e.message];
        }
        res.status(404).send(`
            <!DOCTYPE html>
            <html lang="th">
            <head>
                <meta charset="UTF-8">
                <title>ไม่พบไฟล์หน้าเว็บ</title>
                <style>
                    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0f172a; color: #f8fafc; padding: 40px 20px; display: flex; justify-content: center; align-items: center; min-height: 80vh; margin: 0; }
                    .card { background: #1e293b; border: 1px solid #334155; padding: 32px; border-radius: 16px; max-width: 600px; width: 100%; box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.5); }
                    h1 { color: #f43f5e; margin-top: 0; font-size: 1.5rem; }
                    p { line-height: 1.6; color: #cbd5e1; }
                    .code-box { background: #0b0f19; padding: 14px 18px; border-radius: 8px; font-family: monospace; font-size: 0.9rem; color: #38bdf8; margin: 15px 0; overflow-x: auto; }
                    .badge { background: #22c55e22; color: #4ade80; padding: 4px 10px; border-radius: 999px; font-size: 0.8rem; font-weight: bold; }
                </style>
            </head>
            <body>
                <div class="card">
                    <span class="badge">🚀 เซิร์ฟเวอร์ทำงานสำเร็จแล้ว!</span>
                    <h1>⚠️ แต่ยังไม่พบโฟลเดอร์ public บน GitHub</h1>
                    <p>ระบบ Backend ทำงานได้สมบูรณ์แล้ว แต่ยังไม่มีไฟล์หน้าเว็บ (UI) อัปโหลดขึ้นมาบน GitHub ครับ</p>
                    <p><b>วิธีแก้ไข:</b> ให้เข้าไปที่ Repository บน GitHub แล้วลากโฟลเดอร์ <code>public</code> ทั้งโฟลเดอร์ขึ้นไป</p>
                    <div class="code-box">
                        ไฟล์ที่พบบนเซิร์ฟเวอร์ขณะนี้:<br>
                        ${dirFiles.join(', ')}
                    </div>
                </div>
            </body>
            </html>
        `);
    }
});

// Multer in-memory storage (high speed, no stale files on disk)
const storage = multer.memoryStorage();
const upload = multer({
    storage: storage,
    limits: { fileSize: 100 * 1024 * 1024 } // 100MB
});

// Helper to sanitize filename
function getBaseName(originalName, defaultName = 'document') {
    if (!originalName) return defaultName;
    const ext = path.extname(originalName);
    return path.basename(originalName, ext) || defaultName;
}

// ==========================================
// 1. PDF TO WORD (.docx)
// ==========================================
app.post('/api/pdf/to-word', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' });

        const data = await pdfParse(req.file.buffer);
        const rawText = data.text || '';
        const lines = rawText.split(/\r?\n/);

        const paragraphs = [];
        let currentParagraph = [];

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            if (line.length === 0) {
                if (currentParagraph.length > 0) {
                    paragraphs.push(new docx.Paragraph({
                        children: [new docx.TextRun({ text: currentParagraph.join(' '), size: 24, font: 'Prompt' })],
                        spacing: { after: 180, line: 360 }
                    }));
                    currentParagraph = [];
                }
            } else {
                currentParagraph.push(line);
            }
        }

        if (currentParagraph.length > 0) {
            paragraphs.push(new docx.Paragraph({
                children: [new docx.TextRun({ text: currentParagraph.join(' '), size: 24, font: 'Prompt' })],
                spacing: { after: 180, line: 360 }
            }));
        }

        if (paragraphs.length === 0) {
            paragraphs.push(new docx.Paragraph({
                children: [new docx.TextRun({ text: '(ไม่มีข้อความในไฟล์ PDF หรือเป็นไฟล์ภาพสแกน)', italics: true })]
            }));
        }

        const doc = new docx.Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } // 1 inch
                    }
                },
                children: paragraphs
            }]
        });

        const buffer = await docx.Packer.toBuffer(doc);
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_converted.docx`);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(buffer);
    } catch (err) {
        console.error('PDF to Word Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการแปลงไฟล์ PDF เป็น Word: ' + err.message });
    }
});

// ==========================================
// 2. PDF TO TEXT (.txt)
// ==========================================
app.post('/api/pdf/to-text', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' });

        const data = await pdfParse(req.file.buffer);
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_text.txt`);

        res.setHeader('Content-Type', 'text/plain; charset=utf-8');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(data.text || '');
    } catch (err) {
        console.error('PDF to Text Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการสกัดข้อความ: ' + err.message });
    }
});

// ==========================================
// 3. MERGE PDFs
// ==========================================
app.post('/api/pdf/merge', upload.array('files', 50), async (req, res) => {
    try {
        if (!req.files || req.files.length < 2) {
            return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF อย่างน้อย 2 ไฟล์ขึ้นไป' });
        }

        const mergedPdf = await PDFDocument.create();

        for (const file of req.files) {
            const pdfToMerge = await PDFDocument.load(file.buffer, { ignoreEncryption: true });
            const copiedPages = await mergedPdf.copyPages(pdfToMerge, pdfToMerge.getPageIndices());
            copiedPages.forEach((page) => mergedPdf.addPage(page));
        }

        const pdfBytes = await mergedPdf.save();
        const filename = encodeURIComponent('merged_document.pdf');

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(Buffer.from(pdfBytes));
    } catch (err) {
        console.error('Merge PDF Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการรวม PDF: ' + err.message });
    }
});

// ==========================================
// 4. SPLIT PDF
// ==========================================
app.post('/api/pdf/split', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' });

        const mode = req.body.mode || 'all'; // 'all' or 'range'
        const rangeStr = req.body.range || '';
        const srcPdf = await PDFDocument.load(req.file.buffer, { ignoreEncryption: true });
        const totalPages = srcPdf.getPageCount();
        const zip = new JSZip();
        const baseName = getBaseName(req.file.originalname);

        if (mode === 'all') {
            for (let i = 0; i < totalPages; i++) {
                const newDoc = await PDFDocument.create();
                const [copiedPage] = await newDoc.copyPages(srcPdf, [i]);
                newDoc.addPage(copiedPage);
                const bytes = await newDoc.save();
                zip.file(`${baseName}_page_${i + 1}.pdf`, bytes);
            }
        } else {
            // Parse ranges like "1-3, 5, 7-9"
            const parts = rangeStr.split(',').map(s => s.trim()).filter(Boolean);
            if (parts.length === 0) {
                return res.status(400).json({ error: 'กรุณาระบุช่วงหน้า เช่น 1-3, 5' });
            }

            for (let idx = 0; idx < parts.length; idx++) {
                const part = parts[idx];
                let startPage, endPage;
                if (part.includes('-')) {
                    const [s, e] = part.split('-').map(Number);
                    startPage = Math.max(1, s);
                    endPage = Math.min(totalPages, e);
                } else {
                    const p = Number(part);
                    startPage = p;
                    endPage = p;
                }

                if (isNaN(startPage) || isNaN(endPage) || startPage > totalPages) continue;

                const newDoc = await PDFDocument.create();
                const pageIndices = [];
                for (let p = startPage - 1; p <= endPage - 1; p++) {
                    if (p >= 0 && p < totalPages) pageIndices.push(p);
                }

                if (pageIndices.length > 0) {
                    const copiedPages = await newDoc.copyPages(srcPdf, pageIndices);
                    copiedPages.forEach(page => newDoc.addPage(page));
                    const bytes = await newDoc.save();
                    zip.file(`${baseName}_part_${idx + 1}_(pages_${startPage}-${endPage}).pdf`, bytes);
                }
            }
        }

        const zipBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
        const zipName = encodeURIComponent(`${baseName}_split.zip`);

        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${zipName}`);
        res.send(zipBuffer);
    } catch (err) {
        console.error('Split PDF Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการแยก PDF: ' + err.message });
    }
});

// ==========================================
// 5. ROTATE PDF
// ==========================================
app.post('/api/pdf/rotate', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' });

        const angle = parseInt(req.body.angle) || 90;
        const pagesStr = req.body.pages || 'all';
        const pdfDoc = await PDFDocument.load(req.file.buffer, { ignoreEncryption: true });
        const totalPages = pdfDoc.getPageCount();

        let targetIndices = [];
        if (pagesStr === 'all') {
            targetIndices = pdfDoc.getPageIndices();
        } else {
            const parts = pagesStr.split(',').map(s => s.trim()).filter(Boolean);
            for (const part of parts) {
                if (part.includes('-')) {
                    const [s, e] = part.split('-').map(Number);
                    for (let p = s; p <= e; p++) {
                        if (p >= 1 && p <= totalPages) targetIndices.push(p - 1);
                    }
                } else {
                    const p = Number(part);
                    if (p >= 1 && p <= totalPages) targetIndices.push(p - 1);
                }
            }
        }

        targetIndices = [...new Set(targetIndices)];

        for (const idx of targetIndices) {
            const page = pdfDoc.getPage(idx);
            const currentRotation = page.getRotation().angle;
            page.setRotation(degrees((currentRotation + angle) % 360));
        }

        const pdfBytes = await pdfDoc.save();
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_rotated.pdf`);

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(Buffer.from(pdfBytes));
    } catch (err) {
        console.error('Rotate PDF Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการหมุน PDF: ' + err.message });
    }
});

// ==========================================
// 6. EXTRACT PAGES
// ==========================================
app.post('/api/pdf/extract-pages', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' });

        const pagesStr = req.body.pages || '';
        if (!pagesStr.trim()) return res.status(400).json({ error: 'กรุณาระบุหน้าที่ต้องการดึง เช่น 1, 3, 5-7' });

        const srcDoc = await PDFDocument.load(req.file.buffer, { ignoreEncryption: true });
        const totalPages = srcDoc.getPageCount();

        const targetIndices = [];
        const parts = pagesStr.split(',').map(s => s.trim()).filter(Boolean);
        for (const part of parts) {
            if (part.includes('-')) {
                const [s, e] = part.split('-').map(Number);
                for (let p = s; p <= e; p++) {
                    if (p >= 1 && p <= totalPages) targetIndices.push(p - 1);
                }
            } else {
                const p = Number(part);
                if (p >= 1 && p <= totalPages) targetIndices.push(p - 1);
            }
        }

        const uniqueIndices = [...new Set(targetIndices)].sort((a, b) => a - b);
        if (uniqueIndices.length === 0) {
            return res.status(400).json({ error: 'ไม่พบหมายเลขหน้าที่ระบุในเอกสาร' });
        }

        const newDoc = await PDFDocument.create();
        const copiedPages = await newDoc.copyPages(srcDoc, uniqueIndices);
        copiedPages.forEach(p => newDoc.addPage(p));

        const pdfBytes = await newDoc.save();
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_extracted.pdf`);

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(Buffer.from(pdfBytes));
    } catch (err) {
        console.error('Extract Pages Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการดึงหน้า: ' + err.message });
    }
});

// ==========================================
// 7. DELETE PAGES
// ==========================================
app.post('/api/pdf/delete-pages', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' });

        const pagesToDeleteStr = req.body.pages || '';
        if (!pagesToDeleteStr.trim()) return res.status(400).json({ error: 'กรุณาระบุหน้าที่ต้องการลบ เช่น 2, 4' });

        const srcDoc = await PDFDocument.load(req.file.buffer, { ignoreEncryption: true });
        const totalPages = srcDoc.getPageCount();

        const deleteIndices = new Set();
        const parts = pagesToDeleteStr.split(',').map(s => s.trim()).filter(Boolean);
        for (const part of parts) {
            if (part.includes('-')) {
                const [s, e] = part.split('-').map(Number);
                for (let p = s; p <= e; p++) {
                    if (p >= 1 && p <= totalPages) deleteIndices.add(p - 1);
                }
            } else {
                const p = Number(part);
                if (p >= 1 && p <= totalPages) deleteIndices.add(p - 1);
            }
        }

        const keepIndices = [];
        for (let i = 0; i < totalPages; i++) {
            if (!deleteIndices.has(i)) keepIndices.push(i);
        }

        if (keepIndices.length === 0) {
            return res.status(400).json({ error: 'ไม่สามารถลบทุกหน้าจนหมดเอกสารได้' });
        }

        const newDoc = await PDFDocument.create();
        const copiedPages = await newDoc.copyPages(srcDoc, keepIndices);
        copiedPages.forEach(p => newDoc.addPage(p));

        const pdfBytes = await newDoc.save();
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_cleaned.pdf`);

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(Buffer.from(pdfBytes));
    } catch (err) {
        console.error('Delete Pages Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการลบหน้า: ' + err.message });
    }
});

// ==========================================
// 8. ADD WATERMARK
// ==========================================
app.post('/api/pdf/watermark', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' });

        const text = req.body.text || 'CONFIDENTIAL';
        const opacity = parseFloat(req.body.opacity) || 0.25;
        const pdfDoc = await PDFDocument.load(req.file.buffer, { ignoreEncryption: true });
        const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        const pages = pdfDoc.getPages();

        for (const page of pages) {
            const { width, height } = page.getSize();
            const fontSize = Math.min(width, height) / 8;
            const textWidth = font.widthOfTextAtSize(text, fontSize);
            const textHeight = font.heightAtSize(fontSize);

            page.drawText(text, {
                x: width / 2 - (textWidth / 2) * Math.cos(Math.PI / 4),
                y: height / 2 - (textWidth / 2) * Math.sin(Math.PI / 4),
                size: fontSize,
                font: font,
                color: rgb(0.7, 0.1, 0.1),
                opacity: opacity,
                rotate: degrees(45)
            });
        }

        const pdfBytes = await pdfDoc.save();
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_watermarked.pdf`);

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(Buffer.from(pdfBytes));
    } catch (err) {
        console.error('Watermark Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการเพิ่มลายน้ำ: ' + err.message });
    }
});

// ==========================================
// 9. ADD PAGE NUMBERS
// ==========================================
app.post('/api/pdf/page-numbers', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' });

        const format = req.body.format || 'Page {n} of {total}';
        const pdfDoc = await PDFDocument.load(req.file.buffer, { ignoreEncryption: true });
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const pages = pdfDoc.getPages();
        const total = pages.length;

        for (let i = 0; i < total; i++) {
            const page = pages[i];
            const { width } = page.getSize();
            const text = format.replace('{n}', (i + 1).toString()).replace('{total}', total.toString());
            const fontSize = 10;
            const textWidth = font.widthOfTextAtSize(text, fontSize);

            page.drawText(text, {
                x: (width - textWidth) / 2,
                y: 20,
                size: fontSize,
                font: font,
                color: rgb(0.3, 0.3, 0.3)
            });
        }

        const pdfBytes = await pdfDoc.save();
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_numbered.pdf`);

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(Buffer.from(pdfBytes));
    } catch (err) {
        console.error('Page Numbers Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการใส่เลขหน้า: ' + err.message });
    }
});

// ==========================================
// 10. COMPRESS PDF
// ==========================================
app.post('/api/pdf/compress', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' });

        const originalSize = req.file.buffer.length;
        const pdfDoc = await PDFDocument.load(req.file.buffer, { ignoreEncryption: true });

        // Save with maximum object stream compression
        const compressedBytes = await pdfDoc.save({
            useObjectStreams: true,
            addDefaultPage: false
        });

        const compressedSize = compressedBytes.length;
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_compressed.pdf`);

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.setHeader('X-Original-Size', originalSize.toString());
        res.setHeader('X-Compressed-Size', compressedSize.toString());
        res.send(Buffer.from(compressedBytes));
    } catch (err) {
        console.error('Compress PDF Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการบีบอัด PDF: ' + err.message });
    }
});

// ==========================================
// 11. IMAGES TO PDF
// ==========================================
app.post('/api/image/to-pdf', upload.array('files', 100), async (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'กรุณาอัปโหลดรูปภาพอย่างน้อย 1 รูป' });
        }

        const pageSizeOption = req.body.pageSize || 'a4'; // 'a4' or 'fit'
        const orientation = req.body.orientation || 'portrait'; // 'portrait' or 'landscape'

        const pdfDoc = await PDFDocument.create();

        for (const file of req.files) {
            // Normalize image to PNG or JPEG via sharp
            let imgBuffer = file.buffer;
            let mime = file.mimetype;
            let isPng = mime.includes('png');
            let isJpg = mime.includes('jpeg') || mime.includes('jpg');

            if (!isPng && !isJpg) {
                imgBuffer = await sharp(file.buffer).png().toBuffer();
                isPng = true;
            }

            let embeddedImage;
            try {
                embeddedImage = isPng ? await pdfDoc.embedPng(imgBuffer) : await pdfDoc.embedJpg(imgBuffer);
            } catch {
                // If embedding fails (e.g. invalid colorspace), convert to standard PNG
                imgBuffer = await sharp(file.buffer).toFormat('png').toBuffer();
                embeddedImage = await pdfDoc.embedPng(imgBuffer);
            }

            const imgWidth = embeddedImage.width;
            const imgHeight = embeddedImage.height;

            if (pageSizeOption === 'fit') {
                const page = pdfDoc.addPage([imgWidth, imgHeight]);
                page.drawImage(embeddedImage, {
                    x: 0,
                    y: 0,
                    width: imgWidth,
                    height: imgHeight
                });
            } else {
                // Standard A4: 595.28 x 841.89 points
                let a4W = orientation === 'landscape' ? 841.89 : 595.28;
                let a4H = orientation === 'landscape' ? 595.28 : 841.89;
                const page = pdfDoc.addPage([a4W, a4H]);

                const margin = 20;
                const maxW = a4W - (margin * 2);
                const maxH = a4H - (margin * 2);

                const scale = Math.min(maxW / imgWidth, maxH / imgHeight, 1);
                const drawW = imgWidth * scale;
                const drawH = imgHeight * scale;
                const x = (a4W - drawW) / 2;
                const y = (a4H - drawH) / 2;

                page.drawImage(embeddedImage, {
                    x: x,
                    y: y,
                    width: drawW,
                    height: drawH
                });
            }
        }

        const pdfBytes = await pdfDoc.save();
        const filename = encodeURIComponent('images_converted.pdf');

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(Buffer.from(pdfBytes));
    } catch (err) {
        console.error('Images to PDF Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการแปลงรูปเป็น PDF: ' + err.message });
    }
});

// ==========================================
// 12. IMAGE CONVERT (PNG, JPG, WEBP, AVIF, TIFF, BMP)
// ==========================================
app.post('/api/image/convert', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดรูปภาพ' });

        const targetFormat = (req.body.format || 'png').toLowerCase();
        let pipeline = sharp(req.file.buffer);

        let mimeType = 'image/png';
        let ext = targetFormat;

        switch (targetFormat) {
            case 'jpg':
            case 'jpeg':
                pipeline = pipeline.jpeg({ quality: 90, mozjpeg: true });
                mimeType = 'image/jpeg';
                ext = 'jpg';
                break;
            case 'png':
                pipeline = pipeline.png({ compressionLevel: 8 });
                mimeType = 'image/png';
                break;
            case 'webp':
                pipeline = pipeline.webp({ quality: 85 });
                mimeType = 'image/webp';
                break;
            case 'avif':
                pipeline = pipeline.avif({ quality: 75 });
                mimeType = 'image/avif';
                break;
            case 'tiff':
                pipeline = pipeline.tiff({ compression: 'lzw' });
                mimeType = 'image/tiff';
                break;
            default:
                pipeline = pipeline.png();
                mimeType = 'image/png';
                ext = 'png';
        }

        const outputBuffer = await pipeline.toBuffer();
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}.${ext}`);

        res.setHeader('Content-Type', mimeType);
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(outputBuffer);
    } catch (err) {
        console.error('Image Convert Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการแปลงรูปภาพ: ' + err.message });
    }
});

// ==========================================
// 13. IMAGE RESIZE
// ==========================================
app.post('/api/image/resize', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดรูปภาพ' });

        const width = req.body.width ? parseInt(req.body.width) : null;
        const height = req.body.height ? parseInt(req.body.height) : null;
        const maintainRatio = req.body.maintainRatio !== 'false';

        if (!width && !height) {
            return res.status(400).json({ error: 'กรุณาระบุความกว้างหรือความสูง' });
        }

        let pipeline = sharp(req.file.buffer);
        pipeline = pipeline.resize({
            width: width || undefined,
            height: height || undefined,
            fit: maintainRatio ? 'inside' : 'fill',
            withoutEnlargement: false
        });

        const outputBuffer = await pipeline.toBuffer();
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_resized.png`);

        res.setHeader('Content-Type', 'image/png');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(outputBuffer);
    } catch (err) {
        console.error('Image Resize Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการปรับขนาดรูปภาพ: ' + err.message });
    }
});

// ==========================================
// 14. IMAGE COMPRESS
// ==========================================
app.post('/api/image/compress', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดรูปภาพ' });

        const quality = Math.max(1, Math.min(100, parseInt(req.body.quality) || 70));
        const originalSize = req.file.buffer.length;

        // Auto detect best format or convert to WebP/JPEG for compression
        const format = (req.body.format || 'jpeg').toLowerCase();
        let pipeline = sharp(req.file.buffer);

        let mimeType = 'image/jpeg';
        let ext = 'jpg';

        if (format === 'webp') {
            pipeline = pipeline.webp({ quality: quality });
            mimeType = 'image/webp';
            ext = 'webp';
        } else if (format === 'png') {
            pipeline = pipeline.png({ quality: quality, compressionLevel: 9 });
            mimeType = 'image/png';
            ext = 'png';
        } else {
            pipeline = pipeline.jpeg({ quality: quality, mozjpeg: true });
            mimeType = 'image/jpeg';
            ext = 'jpg';
        }

        const outputBuffer = await pipeline.toBuffer();
        const compressedSize = outputBuffer.length;
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_compressed.${ext}`);

        res.setHeader('Content-Type', mimeType);
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.setHeader('X-Original-Size', originalSize.toString());
        res.setHeader('X-Compressed-Size', compressedSize.toString());
        res.send(outputBuffer);
    } catch (err) {
        console.error('Image Compress Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการบีบอัดรูปภาพ: ' + err.message });
    }
});

// ==========================================
// 15. IMAGE ROTATE & FLIP
// ==========================================
app.post('/api/image/rotate', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดรูปภาพ' });

        const angle = parseInt(req.body.angle) || 0;
        const flipH = req.body.flipH === 'true';
        const flipV = req.body.flipV === 'true';

        let pipeline = sharp(req.file.buffer);

        if (angle) pipeline = pipeline.rotate(angle);
        if (flipH) pipeline = pipeline.flop(); // horizontal
        if (flipV) pipeline = pipeline.flip(); // vertical

        const outputBuffer = await pipeline.toBuffer();
        const filename = encodeURIComponent(`${getBaseName(req.file.originalname)}_transformed.png`);

        res.setHeader('Content-Type', 'image/png');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
        res.send(outputBuffer);
    } catch (err) {
        console.error('Image Rotate Error:', err);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการหมุนรูปภาพ: ' + err.message });
    }
});

// ==========================================
// 16. IMAGE INFO (METADATA)
// ==========================================
app.post('/api/image/info', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'กรุณาอัปโหลดรูปภาพ' });

        const metadata = await sharp(req.file.buffer).metadata();
        res.json({
            width: metadata.width,
            height: metadata.height,
            format: metadata.format,
            space: metadata.space,
            channels: metadata.channels,
            density: metadata.density,
            sizeBytes: req.file.buffer.length
        });
    } catch (err) {
        console.error('Image Info Error:', err);
        res.status(500).json({ error: 'ไม่สามารถอ่านข้อมูลรูปภาพได้: ' + err.message });
    }
});

// Start server
app.listen(PORT, '0.0.0.0', () => {
    console.log(`====================================================`);
    console.log(`🚀 PDF & Image Master กำลังทำงานที่: http://localhost:${PORT}`);
    console.log(`====================================================`);
});
