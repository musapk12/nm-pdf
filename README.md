# NM PDF — Website Starter

Black-and-gold website starter with a working PDF-to-Excel endpoint for PDFs containing selectable text.

## Requirements
- Node.js 18 or newer
- Internet connection for first-time dependency installation

## Run on Windows
1. Install Node.js LTS from https://nodejs.org/
2. Extract `NM_PDF_Starter.zip`.
3. Open the extracted folder in File Explorer.
4. Click the address bar, type `cmd`, and press Enter.
5. Run `npm install`, then `npm start`.
6. Open `http://localhost:3000` in your browser. Keep the command window open.

## Included
- Responsive black-and-gold website
- PDF picker and drag-and-drop
- 20 MB upload limit
- Extracts selectable PDF text into an `.xlsx` workbook
- Basic error handling

## Limitations
This is a starter project, not a published public website. Scanned/image-only PDFs require OCR, which is not included. Extracted text is placed one line per spreadsheet row; complex table layout is not reconstructed. Files are held in server memory for processing and are not intentionally saved to disk by this app.

Before public launch, configure hosting and HTTPS, add rate limits and monitoring, and publish an appropriate privacy notice. The ZIP does not itself publish a website or provide a domain.
