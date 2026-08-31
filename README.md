# FAB Tools Hub

link https://federicomangini83-droid.github.io/FAB-Tools-Hub/

Static GitHub Pages repository containing two utilities:

- **Document Index Creator**: extracts browser-readable text from DOCX and PDF documents and exports JSON, CSV, TXT or ZIP.
- **Table to RAG JSON**: converts Excel/CSV rows to individual JSON chunks inside a ZIP.

## Structure

```text
assets/css/                    Shared visual style
assets/js/                     Shared browser utilities
tools/document-index-creator/  Document tool
tools/table-to-rag-json/       Table conversion tool
backend/                       Original Python implementations
index.html                     Hub homepage
.nojekyll                      GitHub Pages configuration
```

## Publish on GitHub Pages

1. Create a repository named `FAB-Tools-Hub`.
2. Upload the contents of this ZIP to the repository root.
3. In **Settings > Pages**, select **Deploy from a branch**, branch `main`, folder `/ (root)`.
4. Open `https://<username>.github.io/FAB-Tools-Hub/`.

## Processing and privacy

The web tools process files locally in the browser. External CDN libraries are loaded at runtime. The included Python scripts are retained in `backend/` for local use, including the OCR workflow of the document index script.

## Local Python dependencies

```bash
pip install pandas pillow pytesseract python-docx pdfplumber pdf2image openpyxl
```

For OCR, install Tesseract OCR and update the local path in `backend/json_creation_FAB_index.py`.

## GitHub storage

Each tool can save, list, download and delete generated files through the GitHub Contents API. Use a fine-grained personal access token for `FAB-Tools-Hub` with **Contents: Read and write** permission. Storage paths:

- `storage/document-index/`
- `storage/table-to-rag-json/`

The token can remain only in the current page or be stored in browser local storage by selecting the dedicated checkbox. Do not commit tokens into repository files.
