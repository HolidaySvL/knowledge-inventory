# AI Knowledge Inventory — v11 Guided Pre-correction

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

## What is new in v11

v11 improves the pre-correction experience:

- Pre-correction is now placed as a large CTA directly below the diagnosis dashboard.
- Users review the dashboard first, then click **Start Guided Pre-correction**.
- Questions now work like a guided coworker interaction:
  - each question has selectable answer options;
  - the recommended option is clearly marked;
  - a reason is provided for the recommendation;
  - users can still write a custom clarification;
  - users control whether the answer applies to the batch, a segment, selected documents, or one document only.

## Correction scopes

```text
Batch-level
Segment-level
Selected-documents
Document-level
```

## Output

After correction, the app regenerates:

```text
knowledge_inventory_corrected.xlsx
process_map_corrected.md
document_risk_report_corrected.md
correction_ledger.xlsx
knowledge_inventory_corrected_bundle.zip
```

## Core product principle

Minimum human input, maximum knowledge correction.

The system should not ask users to review every document one by one. It should cluster uncertainty, recommend safe choices, leave a custom answer path, and keep every correction traceable.
