# gst-extractor
# GST Invoice Extractor

A local AI-powered GST invoice extraction system built with Python and Flask that processes bulk insurance invoices and generates structured Excel reports automatically.

The application supports `.doc`, `.docx`, and `.pdf` invoices and uses a hybrid architecture combining fast regex-based extraction with local Ollama AI fallback for missing fields.

Unlike cloud AI APIs, this project runs completely offline using Ollama, resulting in:
- zero API cost
- no rate limits
- unlimited invoice processing
- full local privacy

---

## Features

- Bulk ZIP invoice processing
- Supports `.doc`, `.docx`, and `.pdf`
- Automatic extraction of:
  - Bill Number
  - Company Name
  - Invoice Date
  - Grand Total
  - IGST
  - CGST
  - SGST
- Local AI-assisted missing field correction using Ollama
- Fully offline processing
- No API keys required
- Duplicate invoice handling
- GST validation logic
- Structured Excel export
- Browser-based local UI
- Handles 500+ invoices efficiently

---

## Tech Stack

- Python
- Flask
- Pandas
- OpenPyXL
- Python-docx
- PDFPlumber
- Ollama
- LibreOffice

---

## Project Architecture

```text
ZIP Upload
   ↓
Regex Extraction Engine
   ↓
Validation Layer
   ↓
Local Ollama AI Fallback
   ↓
Excel Export