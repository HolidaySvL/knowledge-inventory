# AI Knowledge Inventory — Scope-aware Pre-correction Version

This is a single-file Streamlit deploy version.

Upload these files to GitHub root:

```text
app.py
requirements.txt
packages.txt
runtime.txt
README.md
UPLOAD_TO_GITHUB.md
```

Then deploy on Streamlit Cloud with:

```text
Main file path = app.py
```

## What is new in v10

This version adds **Scope-aware Pre-correction**.

After the initial document scan, users can refine results at different levels:

```text
Batch-level correction
Segment-level correction
Selected-documents correction
Document-level correction
```

The app generates targeted clarification questions, lets the user choose the correction scope, applies the correction, updates the dashboard, and generates a correction ledger.

## Core idea

The system should not ask users to review every document one by one.

Instead, it clusters uncertainty across a document batch and asks a small number of high-impact questions.

Each answer can correct one document, a selected group, or all matching documents.

## Supported upload methods

- Upload folder ZIP
- Upload individual files

## Supported formats

```text
.docx, .doc
.xlsx, .xls, .xlsm
.pptx, .ppt
.pdf
.txt, .md, .csv
.png, .jpg, .jpeg, .webp, .tif, .tiff
.zip
```

## Notes

- `.doc` and `.ppt` conversion requires LibreOffice.
- OCR requires Tesseract.
- If OCR or conversion fails, the app will show extraction warnings instead of crashing.
