# AI Knowledge Intake Agent — v13 Agent Planner

This is a single-file Streamlit deploy version.

Upload these files to the GitHub repo root:

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

## What is new in v13

v13 adds an **Agent Planner** on top of the existing scanner, guided pre-correction, corrected document drafts, and correction ledger.

After the initial scan, the app now shows an Agent Planner control room that can:

- Read the current dashboard state
- Identify whether issues are batch-level, segment-level, or document-level
- Recommend the next best workflow step
- Explain why that path is recommended
- Estimate expected impact
- Prepare guided pre-correction questions for the recommended scope
- Keep the user in control before applying corrections

## Supported workflow

```text
Upload documents
↓
Run initial inventory scan
↓
Review dashboard
↓
Agent Planner recommends next step
↓
Run recommended batch / segment / document-level pre-correction
↓
User selects recommended option or writes custom answer
↓
Apply corrections with controlled scope
↓
Generate corrected documents + correction ledger
↓
Download corrected outputs
```

## Core product idea

This is not a standalone chatbot. It is an agentic knowledge intake layer between messy enterprise documents and AI knowledge systems.

The agent runs the process. Humans confirm business truth. The system keeps the evidence.

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
