const express = require("express");
const multer = require("multer");
const pdfParse = require("pdf-parse");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const isPdf = file.mimetype === "application/pdf" || /\.pdf$/i.test(file.originalname);
    cb(isPdf ? null : new Error("Please upload a PDF file."), isPdf);
  }
});
app.use(express.static(path.join(__dirname, "public")));

app.post("/api/pdf-to-excel", upload.single("pdf"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Please select a PDF file first." });
    const parsed = await pdfParse(req.file.buffer);
    const rawText = (parsed.text || "").replace(/\r/g, "");
    if (!rawText.trim()) {
      return res.status(422).json({ error: "No selectable text found. This may be a scanned PDF; OCR is not included in this starter." });
    }
    const lines = rawText.split("\n").map(line => line.trim()).filter(Boolean);
    const worksheet = XLSX.utils.aoa_to_sheet([["Extracted PDF Text"], ...lines.map(line => [line])]);
    worksheet["!cols"] = [{ wch: 100 }];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "PDF Text");
    const output = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
    const safeName = path.parse(req.file.originalname).name.replace(/[^\w-]+/g, "_") || "converted";
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName}.xlsx"`);
    res.send(output);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Conversion failed. Please try another PDF." });
  }
});

app.use((err, _req, res, _next) => {
  if (err instanceof multer.MulterError && err.code === "LIMIT_FILE_SIZE") {
    return res.status(413).json({ error: "File is too large. Maximum size is 20 MB." });
  }
  res.status(400).json({ error: err.message || "Request error." });
});
app.listen(PORT, () => console.log(`NM PDF is running at http://localhost:${PORT}`));
