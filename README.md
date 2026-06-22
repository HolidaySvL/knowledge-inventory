# AI Knowledge Inventory — v12 Corrected Documents

This is the single-file Streamlit deploy version.

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

## What is new in v12

v12 adds **Corrected Document Export** after guided pre-correction.

After users answer the clarification questions and apply corrections, the app now generates:

```text
corrected_documents.zip
  corrected_documents/*.docx
  corrected_documents/*.md
  corrected_documents_manifest.xlsx
```

The app also keeps the governance outputs:

```text
knowledge_inventory_corrected.xlsx
document_risk_report_corrected.md
process_map_corrected.md
correction_ledger.xlsx
```

## User flow

```text
Upload documents
Run initial diagnosis
Review dashboard
Start Guided Pre-correction
Choose correction scope
Answer Claude-style recommended options or write a custom answer
Apply corrections
Review Correction Ledger
Download corrected documents
```

## Important note

The app does not overwrite the original uploaded files. It generates corrected draft documents as new downloadable outputs.

This is safer for enterprise governance because the original source file remains untouched, while the correction ledger records what changed and why.
