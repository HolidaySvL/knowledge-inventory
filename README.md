# AI Knowledge Inventory - Web Upload Version

This is a single-file Streamlit deploy version.

Upload these files to GitHub root:

```text
app.py
requirements.txt
packages.txt
runtime.txt
README.md
```

Then deploy on Streamlit Cloud with:

```text
Main file path = app.py
```

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
