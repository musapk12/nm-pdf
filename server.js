const express = require("express");
const multer = require("multer");
const pdfParse = require("pdf-parse");
const XLSX = require("xlsx");
const { PDFDocument, degrees } = require("pdf-lib");
const PDFKit = require("pdfkit");
const JSZip = require("jszip");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024, files: 20 }
});

app.use(express.static(path.join(__dirname, "public")));

function safeName(name) {
  return path.parse(name || "converted").name.replace(/[^\w-]+/g, "_") || "converted";
}
function requirePdf(file) {
  return file && (file.mimetype === "application/pdf" || /\.pdf$/i.test(file.originalname));
}
function requireExcel(file) {
  return file && (/\.(xlsx|xls|csv)$/i.test(file.originalname) ||
    /spreadsheet|excel|csv/i.test(file.mimetype || ""));
}
function sendError(res, err) {
  console.error(err);
  res.status(400).json({ error: err.message || "The operation failed. Please check your file and try again." });
}

app.get("/api/health", (_req, res) => res.json({ ok: true, app: "NM PDF" }));

// PDF to Excel: extracts selectable PDF text into rows in an Excel workbook.
app.post("/api/pdf-to-excel", upload.single("pdf"), async (req, res) => {
  try {
    if (!req.file || !requirePdf(req.file)) return res.status(400).json({ error: "Please select a PDF file." });
    const parsed = await pdfParse(req.file.buffer);
    const rawText = (parsed.text || "").replace(/\r/g, "");
    if (!rawText.trim()) return res.status(422).json({ error: "No selectable text found. Scanned PDFs need OCR, which is not included." });
    const lines = rawText.split("\n").map(s => s.trim()).filter(Boolean);
    const ws = XLSX.utils.aoa_to_sheet([["Extracted PDF Text"], ...lines.map(line => [line])]);
    ws["!cols"] = [{ wch: 100 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "PDF Text");
    const output = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName(req.file.originalname)}.xlsx"`);
    res.send(output);
  } catch (err) { sendError(res, err); }
});

// PDF to Word: puts extracted selectable text into an editable DOCX file.
app.post("/api/pdf-to-word", upload.single("pdf"), async (req, res) => {
  try {
    if (!req.file || !requirePdf(req.file)) return res.status(400).json({ error: "Please select a PDF file." });
    const parsed = await pdfParse(req.file.buffer);
    const rawText = (parsed.text || "").trim();
    if (!rawText) return res.status(422).json({ error: "No selectable text found. Scanned PDFs need OCR, which is not included." });
    const paragraphs = rawText.split(/\n+/).map(line => new Paragraph({ children: [new TextRun(line)] }));
    const doc = new Document({ sections: [{ children: paragraphs }] });
    const output = await Packer.toBuffer(doc);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName(req.file.originalname)}.docx"`);
    res.send(output);
  } catch (err) { sendError(res, err); }
});

// Merge multiple PDFs into one.
app.post("/api/merge-pdf", upload.array("pdfs", 20), async (req, res) => {
  try {
    const files = req.files || [];
    if (files.length < 2) return res.status(400).json({ error: "Select at least two PDF files to merge." });
    if (files.some(f => !requirePdf(f))) return res.status(400).json({ error: "All selected files must be PDFs." });
    const merged = await PDFDocument.create();
    for (const file of files) {
      const source = await PDFDocument.load(file.buffer);
      const pages = await merged.copyPages(source, source.getPageIndices());
      pages.forEach(page => merged.addPage(page));
    }
    const output = Buffer.from(await merged.save());
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", 'attachment; filename="nm-pdf-merged.pdf"');
    res.send(output);
  } catch (err) { sendError(res, err); }
});

// Split PDF into one-page PDF files, returned in a ZIP.
app.post("/api/split-pdf", upload.single("pdf"), async (req, res) => {
  try {
    if (!req.file || !requirePdf(req.file)) return res.status(400).json({ error: "Please select a PDF file." });
    const source = await PDFDocument.load(req.file.buffer);
    if (source.getPageCount() < 2) return res.status(400).json({ error: "This PDF has only one page, so it cannot be split." });
    const zip = new JSZip();
    for (let i = 0; i < source.getPageCount(); i++) {
      const onePage = await PDFDocument.create();
      const [page] = await onePage.copyPages(source, [i]);
      onePage.addPage(page);
      zip.file(`page-${String(i + 1).padStart(3, "0")}.pdf`, await onePage.save());
    }
    const output = await zip.generateAsync({ type: "nodebuffer" });
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName(req.file.originalname)}-pages.zip"`);
    res.send(output);
  } catch (err) { sendError(res, err); }
});

// Excel to PDF: renders worksheet cell values as a readable table/text PDF.
// Note: this is a basic data export, not a pixel-perfect Microsoft Excel print layout.
app.post("/api/excel-to-pdf", upload.single("excel"), async (req, res) => {
  try {
    if (!req.file || !requireExcel(req.file)) return res.status(400).json({ error: "Please select an Excel (.xlsx, .xls) or CSV file." });
    const wb = XLSX.read(req.file.buffer, { type: "buffer" });
    const firstSheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false, defval: "" });
    if (!rows.length) return res.status(400).json({ error: "The selected spreadsheet is empty." });

    const doc = new PDFKit({ size: "A4", layout: "landscape", margin: 35, bufferPages: true });
    const chunks = [];
    doc.on("data", chunk => chunks.push(chunk));
    const done = new Promise((resolve, reject) => { doc.on("end", resolve); doc.on("error", reject); });
    doc.fontSize(14).text("NM PDF - Excel to PDF", { align: "center" });
    doc.moveDown();
    const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const maxCols = Math.max(1, ...rows.map(r => r.length));
    const colWidth = Math.max(35, Math.min(150, pageWidth / maxCols));
    const rowHeight = 20;
    rows.forEach((row, ri) => {
      if (doc.y + rowHeight > doc.page.height - doc.page.margins.bottom) doc.addPage({ size: "A4", layout: "landscape", margin: 35 });
      const y = doc.y;
      for (let ci = 0; ci < maxCols; ci++) {
        const x = doc.page.margins.left + ci * colWidth;
        const value = String(row[ci] == null ? "" : row[ci]).replace(/\s+/g, " ");
        doc.rect(x, y, colWidth, rowHeight).stroke();
        doc.fontSize(7).text(value.slice(0, 120), x + 3, y + 5, { width: colWidth - 6, height: rowHeight - 4, ellipsis: true });
      }
      doc.y = y + rowHeight;
    });
    doc.end();
    await done;
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName(req.file.originalname)}.pdf"`);
    res.send(Buffer.concat(chunks));
  } catch (err) { sendError(res, err); }
});

// Rotate every page by the requested angle (90, 180, or 270).
app.post("/api/rotate-pdf", upload.single("pdf"), async (req, res) => {
  try {
    if (!req.file || !requirePdf(req.file)) return res.status(400).json({ error: "Please select a PDF file." });
    const angle = Number(req.body.angle || 90);
    if (![90, 180, 270].includes(angle)) return res.status(400).json({ error: "Choose 90, 180, or 270 degrees." });
    const pdf = await PDFDocument.load(req.file.buffer);
    pdf.getPages().forEach(page => page.setRotation(degrees((page.getRotation().angle + angle) % 360)));
    const output = Buffer.from(await pdf.save());
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName(req.file.originalname)}-rotated.pdf"`);
    res.send(output);
  } catch (err) { sendError(res, err); }
});

app.use((err, _req, res, _next) => {
  if (err instanceof multer.MulterError && err.code === "LIMIT_FILE_SIZE") {
    return res.status(413).json({ error: "File too large. Maximum size is 20 MB per file." });
  }
  res.status(400).json({ error: err.message || "Request error." });
});

app.listen(PORT, () => console.log(`NM PDF is running on port ${PORT}`));
