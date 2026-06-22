from pathlib import Path
from io import BytesIO
from typing import Dict, Any, List, Tuple
import zipfile
import tempfile
import subprocess
import shutil
import os
import json
import re

import pandas as pd
import streamlit as st
import requests
import plotly.express as px
from dotenv import load_dotenv

from docx import Document
from pptx import Presentation
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

load_dotenv()

# ============================================================
# Config
# ============================================================

SUPPORTED_EXTENSIONS = {
    ".docx", ".doc",
    ".xlsx", ".xls", ".xlsm",
    ".pptx", ".ppt",
    ".pdf",
    ".txt", ".md", ".csv",
    ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff",
}
UPLOAD_EXTENSIONS = SUPPORTED_EXTENSIONS | {".zip"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"}

REQUIRED_FIELDS = [
    "document_name",
    "source_path",
    "file_type",
    "document_type",
    "department",
    "business_process",
    "knowledge_category",
    "summary",
    "risk_level",
    "risk_reason",
    "recommended_action",
    "can_enter_kb",
    "needs_human_review"
]

CLASSIFY_PROMPT = """
You are an enterprise AI knowledge inventory assistant.

Your task is to analyze one business document and produce a structured inventory record.

Important rules:
- Do not rewrite the document.
- Do not invent facts.
- If something is unclear, use "Unknown" or "Needs Review".
- The output must be valid JSON only.
- Risk level must be one of: Low, Medium, High.
- can_enter_kb must be one of: Yes, No, Needs Review.
- needs_human_review must be true or false.

Metadata:
{{metadata_json}}

Document text:
{{document_text}}

Return this JSON structure exactly:
{
  "document_name": "",
  "source_path": "",
  "file_type": "",
  "document_type": "",
  "department": "",
  "business_process": "",
  "knowledge_category": "",
  "summary": "",
  "risk_level": "Low / Medium / High",
  "risk_reason": "",
  "recommended_action": "",
  "can_enter_kb": "Yes / No / Needs Review",
  "needs_human_review": true
}
"""

PROCESS_MAP_PROMPT = """
You are an enterprise knowledge management consultant.

Based on the following knowledge inventory JSON, generate a concise process map in Markdown.

Focus on:
- departments
- business processes
- related documents
- documentation gaps
- which processes are ready for knowledge-base preparation
- which processes need cleanup or human review

Inventory JSON:
{{inventory_json}}
"""

RISK_REPORT_PROMPT = """
You are an enterprise document governance analyst.

Based on the following knowledge inventory JSON, generate a concise document risk report in Markdown.

Include:
1. Executive summary
2. High-risk documents
3. Medium-risk documents
4. Documents needing human review
5. Recommended cleanup priorities
6. Priority candidates for knowledge-base preparation

Inventory JSON:
{{inventory_json}}
"""

# ============================================================
# File reading
# ============================================================

def safe_decode(file_bytes: bytes) -> str:
    for enc in ["utf-8", "utf-8-sig", "latin-1"]:
        try:
            return file_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return file_bytes.decode("latin-1", errors="ignore")

def ocr_image_bytes(file_bytes: bytes) -> str:
    try:
        from PIL import Image
        import pytesseract
        img = Image.open(BytesIO(file_bytes))
        text = pytesseract.image_to_string(img)
        return text.strip()
    except Exception as exc:
        return f"[Image detected, but OCR is unavailable or failed: {exc}]"

def extract_docx_bytes(file_bytes: bytes, ocr_images: bool = False) -> str:
    doc = Document(BytesIO(file_bytes))
    parts = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            parts.append(text)

    for table_idx, table in enumerate(doc.tables, start=1):
        parts.append(f"\n[Table {table_idx}]")
        for row in table.rows:
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))

    if ocr_images:
        try:
            with zipfile.ZipFile(BytesIO(file_bytes)) as z:
                media = [n for n in z.namelist() if n.startswith("word/media/")]
                for idx, name in enumerate(media, start=1):
                    suffix = Path(name).suffix.lower()
                    if suffix in IMAGE_EXTENSIONS:
                        parts.append(f"\n[Embedded Image {idx}: {Path(name).name}]")
                        parts.append(ocr_image_bytes(z.read(name)))
        except Exception as exc:
            parts.append(f"\n[Embedded image extraction failed: {exc}]")

    return "\n".join(parts)

def extract_pptx_bytes(file_bytes: bytes, ocr_images: bool = False) -> str:
    prs = Presentation(BytesIO(file_bytes))
    parts = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        parts.append(f"\n[Slide {slide_idx}]")
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                text = shape.text.strip()
                if text:
                    parts.append(text)

    if ocr_images:
        try:
            with zipfile.ZipFile(BytesIO(file_bytes)) as z:
                media = [n for n in z.namelist() if n.startswith("ppt/media/")]
                for idx, name in enumerate(media, start=1):
                    suffix = Path(name).suffix.lower()
                    if suffix in IMAGE_EXTENSIONS:
                        parts.append(f"\n[Embedded Image {idx}: {Path(name).name}]")
                        parts.append(ocr_image_bytes(z.read(name)))
        except Exception as exc:
            parts.append(f"\n[Embedded image extraction failed: {exc}]")

    return "\n".join(parts)

def extract_excel_bytes(file_bytes: bytes, suffix: str) -> Tuple[str, List[str]]:
    parts = []
    sheet_names = []

    if suffix == ".csv":
        try:
            df = pd.read_csv(BytesIO(file_bytes), dtype=str).fillna("")
            parts.append(df.head(200).to_csv(index=False))
            return "\n".join(parts), ["CSV"]
        except Exception as exc:
            return f"[CSV file detected but could not be read: {exc}]", []

    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as exc:
        return f"[Excel file detected but could not be read: {exc}]", []

    sheet_names = xls.sheet_names
    for sheet in sheet_names:
        try:
            df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, dtype=str)
            df = df.fillna("")
            parts.append(f"\n[Sheet: {sheet}]")
            if df.empty:
                parts.append("[Empty sheet]")
            else:
                parts.append(df.head(120).to_csv(index=False))
        except Exception as exc:
            parts.append(f"\n[Sheet: {sheet}]")
            parts.append(f"[Could not read sheet: {exc}]")

    return "\n".join(parts), sheet_names

def extract_pdf_bytes(file_bytes: bytes, ocr_images: bool = False, max_ocr_pages: int = 5) -> str:
    try:
        import fitz
    except Exception as exc:
        return f"[PDF detected but PyMuPDF is unavailable: {exc}]"

    parts = []
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    for page_idx, page in enumerate(doc, start=1):
        text = page.get_text("text").strip()
        parts.append(f"\n[PDF Page {page_idx}]")
        if text:
            parts.append(text)
        else:
            parts.append("[No embedded text detected on this page.]")

        if ocr_images and page_idx <= max_ocr_pages and len(text) < 80:
            try:
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                img_bytes = pix.tobytes("png")
                parts.append("[OCR attempt]")
                parts.append(ocr_image_bytes(img_bytes))
            except Exception as exc:
                parts.append(f"[PDF OCR failed: {exc}]")

    return "\n".join(parts)

def convert_legacy_office_bytes(file_bytes: bytes, original_name: str, target_ext: str) -> Tuple[bytes, str]:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice is not installed or not found in PATH. Cannot convert legacy Office file.")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        input_path = tmp / original_name
        input_path.write_bytes(file_bytes)

        convert_to = "docx" if target_ext == ".docx" else "pptx"
        subprocess.run(
            [soffice, "--headless", "--convert-to", convert_to, "--outdir", str(tmp), str(input_path)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        converted = input_path.with_suffix(target_ext)
        if not converted.exists():
            candidates = list(tmp.glob(f"*{target_ext}"))
            if not candidates:
                raise RuntimeError("LibreOffice conversion finished but no converted file was found.")
            converted = candidates[0]

        return converted.read_bytes(), converted.name

def extract_text_from_bytes(name: str, file_bytes: bytes, source_path: str, ocr_images: bool = False) -> Dict[str, Any]:
    suffix = Path(name).suffix.lower()
    warnings = []
    sheet_names = []

    try:
        if suffix == ".docx":
            text = extract_docx_bytes(file_bytes, ocr_images=ocr_images)

        elif suffix == ".doc":
            try:
                converted_bytes, converted_name = convert_legacy_office_bytes(file_bytes, Path(name).name, ".docx")
                text = extract_docx_bytes(converted_bytes, ocr_images=ocr_images)
                warnings.append(f"Legacy .doc converted through LibreOffice to {converted_name}.")
            except Exception as exc:
                text = f"[Legacy .doc file detected but conversion failed: {exc}]"
                warnings.append(str(exc))

        elif suffix in [".xlsx", ".xls", ".xlsm", ".csv"]:
            text, sheet_names = extract_excel_bytes(file_bytes, suffix)

        elif suffix == ".pptx":
            text = extract_pptx_bytes(file_bytes, ocr_images=ocr_images)

        elif suffix == ".ppt":
            try:
                converted_bytes, converted_name = convert_legacy_office_bytes(file_bytes, Path(name).name, ".pptx")
                text = extract_pptx_bytes(converted_bytes, ocr_images=ocr_images)
                warnings.append(f"Legacy .ppt converted through LibreOffice to {converted_name}.")
            except Exception as exc:
                text = f"[Legacy .ppt file detected but conversion failed: {exc}]"
                warnings.append(str(exc))

        elif suffix == ".pdf":
            text = extract_pdf_bytes(file_bytes, ocr_images=ocr_images)

        elif suffix in IMAGE_EXTENSIONS:
            text = ocr_image_bytes(file_bytes) if ocr_images else "[Image file detected. Enable OCR to extract text.]"

        elif suffix in [".txt", ".md"]:
            text = safe_decode(file_bytes)

        else:
            text = f"[Unsupported file type: {suffix}]"

    except Exception as exc:
        text = f"[Failed to extract text: {exc}]"
        warnings.append(str(exc))

    return {
        "document_name": Path(name).name,
        "source_path": source_path,
        "text": text,
        "file_type": suffix,
        "sheet_names": sheet_names,
        "extraction_warnings": warnings,
    }

def extract_text_from_upload(uploaded_file, ocr_images: bool = False) -> Dict[str, Any]:
    return extract_text_from_bytes(uploaded_file.name, uploaded_file.getvalue(), "uploaded_file", ocr_images=ocr_images)

def extract_docs_from_zip_upload(uploaded_file, ocr_images: bool = False) -> List[Dict[str, Any]]:
    docs = []
    file_bytes = uploaded_file.getvalue()

    with zipfile.ZipFile(BytesIO(file_bytes)) as z:
        for member in z.infolist():
            if member.is_dir():
                continue

            member_name = member.filename
            path = Path(member_name)
            suffix = path.suffix.lower()

            if suffix not in SUPPORTED_EXTENSIONS:
                continue
            if path.name.startswith("~$") or "__MACOSX" in path.parts:
                continue

            with z.open(member) as f:
                content = f.read()

            docs.append(
                extract_text_from_bytes(
                    name=member_name,
                    file_bytes=content,
                    source_path=f"{uploaded_file.name}/{member_name}",
                    ocr_images=ocr_images,
                )
            )

    return docs

# ============================================================
# AI Client
# ============================================================

def get_default_ai_config() -> Dict[str, Any]:
    return {
        "mode": os.getenv("AI_MODE", "heuristic"),
        "api_key": os.getenv("AI_API_KEY", ""),
        "base_url": os.getenv("AI_BASE_URL", ""),
        "model": os.getenv("AI_MODEL", ""),
    }

def safe_json_parse(text: str) -> Dict[str, Any]:
    if not text:
        raise ValueError("Empty AI response.")

    cleaned = text.strip()

    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```[a-zA-Z]*\n?", "", cleaned)
        cleaned = re.sub(r"\n?```$", "", cleaned)

    try:
        return json.loads(cleaned)
    except Exception:
        pass

    match = re.search(r"\{[\s\S]*\}", cleaned)
    if match:
        return json.loads(match.group(0))

    raise ValueError("Could not parse JSON from AI response.")

def normalize_result(data: Dict[str, Any], metadata: Dict[str, Any]) -> Dict[str, Any]:
    result = {}
    for field in REQUIRED_FIELDS:
        result[field] = data.get(field, "")

    result["document_name"] = result["document_name"] or metadata.get("document_name", "")
    result["source_path"] = result["source_path"] or metadata.get("source_path", "")
    result["file_type"] = result["file_type"] or metadata.get("file_type", "")

    if isinstance(result["needs_human_review"], str):
        result["needs_human_review"] = result["needs_human_review"].strip().lower() in ["true", "yes", "y", "1"]

    if result["risk_level"] not in ["Low", "Medium", "High"]:
        result["risk_level"] = "Medium"

    if result["can_enter_kb"] not in ["Yes", "No", "Needs Review"]:
        result["can_enter_kb"] = "Needs Review"

    return result

def heuristic_analyze(doc_text: str, metadata: Dict[str, Any]) -> Dict[str, Any]:
    name = metadata.get("document_name", "")
    lower = (name + "\n" + doc_text[:4000]).lower()

    document_type = "Unknown"
    department = "Unknown"
    business_process = "Unknown"
    knowledge_category = "General Business Knowledge"

    if "supplier" in lower or "vendor" in lower:
        department = "Procurement"
        business_process = "Supplier / Vendor Management"
    if "purchase" in lower or "procurement" in lower:
        department = "Procurement"
        business_process = "Purchase Approval"
    if "quality" in lower or "inspection" in lower or "complaint" in lower:
        department = "Quality"
        business_process = "Quality Management"
    if "maintenance" in lower or "equipment" in lower or "machine" in lower:
        department = "Maintenance"
        business_process = "Equipment Maintenance"
    if "production" in lower or "abnormal" in lower:
        department = "Production"
        business_process = "Production Issue Handling"
    if "safety" in lower or "ppe" in lower or "emergency" in lower:
        department = "EHS / Safety"
        business_process = "Safety Operation"
    if "training" in lower or "new employee" in lower or "onboarding" in lower:
        department = "HR / Operations"
        business_process = "Employee Onboarding"

    suffix = metadata.get("file_type", "")
    if "sop" in lower or "procedure" in lower or "process" in lower:
        document_type = "SOP / Process Document"
        knowledge_category = "Process Knowledge"
    elif suffix in [".xlsx", ".xls", ".xlsm", ".csv"]:
        document_type = "Spreadsheet / Operational Record"
        knowledge_category = "Operational Record / Checklist"
    elif suffix in [".pdf"]:
        document_type = "PDF Document"
        knowledge_category = "Documented Business Knowledge"
    elif suffix in [".ppt", ".pptx"]:
        document_type = "Presentation / Training Material"
        knowledge_category = "Training / Proposal Knowledge"
    elif suffix in IMAGE_EXTENSIONS:
        document_type = "Image / Screenshot"
        knowledge_category = "Visual Reference"
    elif "meeting" in lower or "action item" in lower:
        document_type = "Meeting Notes"
        knowledge_category = "Decision / Action Item Record"
    elif "manual" in lower or "guide" in lower:
        document_type = "Manual / Guide"
        knowledge_category = "Training / Reference Knowledge"

    risk_reasons = []
    if "version" not in lower and " v1" not in lower and " v2" not in lower and "_v1" not in lower and "_v2" not in lower:
        risk_reasons.append("Version is not clearly identified.")
    if "owner" not in lower and "responsible" not in lower and "approved by" not in lower:
        risk_reasons.append("Document owner or responsible role may be unclear.")
    if "unknown" in lower or "tbd" in lower or "to be confirmed" in lower or "needs review" in lower:
        risk_reasons.append("Contains unresolved or unclear information.")
    if "may" in lower or "should" in lower or "if needed" in lower or "as appropriate" in lower:
        risk_reasons.append("Some rules appear conditional or vague.")
    if "[failed" in lower or "[legacy" in lower or "ocr is unavailable" in lower:
        risk_reasons.append("Text extraction may be incomplete.")

    risk_level = "Low"
    if len(risk_reasons) >= 2:
        risk_level = "High"
    elif len(risk_reasons) == 1:
        risk_level = "Medium"

    summary = doc_text.strip().replace("\n", " ")[:450] or "No readable text extracted."
    can_enter_kb = "Yes" if risk_level == "Low" else "Needs Review"

    return normalize_result({
        "document_name": name,
        "source_path": metadata.get("source_path", ""),
        "file_type": metadata.get("file_type", ""),
        "document_type": document_type,
        "department": department,
        "business_process": business_process,
        "knowledge_category": knowledge_category,
        "summary": summary,
        "risk_level": risk_level,
        "risk_reason": " ".join(risk_reasons) if risk_reasons else "No major risk detected by fallback scan.",
        "recommended_action": "Review metadata and confirm whether this document is the current source of truth." if risk_level != "Low" else "Can be considered for knowledge base preparation.",
        "can_enter_kb": can_enter_kb,
        "needs_human_review": risk_level != "Low"
    }, metadata)

def call_gemini(prompt: str, config: Dict[str, Any]) -> str:
    api_key = config.get("api_key", "")
    model = config.get("model") or "gemini-2.5-flash"
    if not api_key:
        raise RuntimeError("API key is empty.")

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1}
    }

    response = requests.post(url, params={"key": api_key}, json=payload, timeout=90)
    response.raise_for_status()
    data = response.json()
    return data["candidates"][0]["content"]["parts"][0]["text"]

def call_openai_compatible(prompt: str, config: Dict[str, Any]) -> str:
    api_key = config.get("api_key", "")
    base_url = (config.get("base_url") or "").rstrip("/")
    model = config.get("model", "")

    if not api_key:
        raise RuntimeError("API key is empty.")
    if not base_url:
        raise RuntimeError("Base URL is empty.")
    if not model:
        raise RuntimeError("Model is empty.")

    url = base_url if base_url.endswith("/chat/completions") else f"{base_url}/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": "You are a precise enterprise knowledge inventory assistant. Follow output format strictly."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.1
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    response = requests.post(url, headers=headers, json=payload, timeout=90)
    response.raise_for_status()
    data = response.json()
    return data["choices"][0]["message"]["content"]

def call_anthropic(prompt: str, config: Dict[str, Any]) -> str:
    api_key = config.get("api_key", "")
    base_url = (config.get("base_url") or "https://api.anthropic.com").rstrip("/")
    model = config.get("model") or "claude-3-5-haiku-latest"

    if not api_key:
        raise RuntimeError("API key is empty.")

    url = f"{base_url}/v1/messages"
    payload = {
        "model": model,
        "max_tokens": 4096,
        "temperature": 0.1,
        "messages": [{"role": "user", "content": prompt}]
    }
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, json=payload, timeout=90)
    response.raise_for_status()
    data = response.json()
    return data["content"][0]["text"]

def call_llm(prompt: str, config: Dict[str, Any]) -> str:
    mode = config.get("mode", "heuristic")
    if mode == "gemini":
        return call_gemini(prompt, config)
    if mode == "anthropic":
        return call_anthropic(prompt, config)
    if mode == "openai_compatible":
        return call_openai_compatible(prompt, config)
    raise RuntimeError("No API mode selected.")

def analyze_document(doc_text: str, metadata: Dict[str, Any], max_chars: int, ai_config: Dict[str, Any]) -> Dict[str, Any]:
    if ai_config.get("mode") == "heuristic":
        return heuristic_analyze(doc_text, metadata)

    prompt = (
        CLASSIFY_PROMPT
        .replace("{{metadata_json}}", json.dumps(metadata, ensure_ascii=False, indent=2))
        .replace("{{document_text}}", doc_text[:max_chars])
    )

    try:
        raw = call_llm(prompt, ai_config)
        data = safe_json_parse(raw)
        return normalize_result(data, metadata)
    except Exception as exc:
        fallback = heuristic_analyze(doc_text, metadata)
        fallback["risk_level"] = "High"
        fallback["risk_reason"] = f"AI call failed or returned invalid JSON. Fallback used. Error: {exc}"
        fallback["recommended_action"] = "Check endpoint, API key, model, or output format, then rerun analysis."
        fallback["can_enter_kb"] = "Needs Review"
        fallback["needs_human_review"] = True
        return fallback

def heuristic_process_map(inventory: List[Dict[str, Any]]) -> str:
    groups = {}
    for item in inventory:
        dept = item.get("department") or "Unknown"
        proc = item.get("business_process") or "Unknown"
        groups.setdefault(dept, {}).setdefault(proc, []).append(item)

    lines = ["# Process Map", "", "Generated from the knowledge inventory.", ""]
    for dept, processes in groups.items():
        lines.append(f"## {dept}\n")
        for proc, docs in processes.items():
            lines.append(f"### {proc}\n")
            lines.append("Related documents:")
            for doc in docs:
                lines.append(f"- {doc.get('document_name')} — {doc.get('document_type')}")
            needs_review = [d for d in docs if d.get("needs_human_review")]
            lines.append(f"\nDocumentation status: {'Needs review' if needs_review else 'Generally usable'}")
            if needs_review:
                lines.append("Risk notes:")
                for doc in needs_review[:5]:
                    lines.append(f"- {doc.get('document_name')}: {doc.get('risk_reason')}")
            lines.append("")
    return "\n".join(lines)

def heuristic_risk_report(inventory: List[Dict[str, Any]]) -> str:
    high = [x for x in inventory if x.get("risk_level") == "High"]
    medium = [x for x in inventory if x.get("risk_level") == "Medium"]
    kb_ready = [x for x in inventory if x.get("can_enter_kb") == "Yes"]
    review = [x for x in inventory if x.get("needs_human_review")]

    lines = ["# Document Risk Report", "", "## Executive Summary", ""]
    lines.append(f"- Total documents analyzed: {len(inventory)}")
    lines.append(f"- High-risk documents: {len(high)}")
    lines.append(f"- Medium-risk documents: {len(medium)}")
    lines.append(f"- Documents needing human review: {len(review)}")
    lines.append(f"- Documents potentially ready for KB preparation: {len(kb_ready)}\n")

    lines.append("## Priority Review Items\n")
    if review:
        for doc in review:
            lines.append(f"### {doc.get('document_name')}")
            lines.append(f"- Department: {doc.get('department')}")
            lines.append(f"- Process: {doc.get('business_process')}")
            lines.append(f"- Risk level: {doc.get('risk_level')}")
            lines.append(f"- Risk reason: {doc.get('risk_reason')}")
            lines.append(f"- Recommended action: {doc.get('recommended_action')}\n")
    else:
        lines.append("No documents were flagged for human review.\n")

    lines.append("## Priority KB Candidates\n")
    if kb_ready:
        for doc in kb_ready:
            lines.append(f"- {doc.get('document_name')} — {doc.get('business_process')}")
    else:
        lines.append("No documents were marked as ready without review.")
    return "\n".join(lines)

def generate_process_map(inventory: List[Dict[str, Any]], ai_config: Dict[str, Any]) -> str:
    if ai_config.get("mode") == "heuristic":
        return heuristic_process_map(inventory)
    prompt = PROCESS_MAP_PROMPT.replace("{{inventory_json}}", json.dumps(inventory, ensure_ascii=False, indent=2))
    try:
        return call_llm(prompt, ai_config)
    except Exception as exc:
        return heuristic_process_map(inventory) + f"\n\n> AI generation failed; fallback used. Error: {exc}\n"

def generate_risk_report(inventory: List[Dict[str, Any]], ai_config: Dict[str, Any]) -> str:
    if ai_config.get("mode") == "heuristic":
        return heuristic_risk_report(inventory)
    prompt = RISK_REPORT_PROMPT.replace("{{inventory_json}}", json.dumps(inventory, ensure_ascii=False, indent=2))
    try:
        return call_llm(prompt, ai_config)
    except Exception as exc:
        return heuristic_risk_report(inventory) + f"\n\n> AI generation failed; fallback used. Error: {exc}\n"

# ============================================================
# Export
# ============================================================

def inventory_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Knowledge Inventory")
        ws = writer.book["Knowledge Inventory"]

        header_fill = PatternFill("solid", fgColor="D9EAF7")
        header_font = Font(bold=True)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        for col_idx, col in enumerate(ws.columns, start=1):
            max_len = 0
            for cell in col:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, min(len(value), 60))
            ws.column_dimensions[get_column_letter(col_idx)].width = max(12, max_len + 2)

        ws.freeze_panes = "A2"

    return buffer.getvalue()

def markdown_bytes(text: str) -> bytes:
    return (text or "").encode("utf-8")

def report_zip_bytes(df: pd.DataFrame, process_map: str, risk_report: str) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("knowledge_inventory.xlsx", inventory_excel_bytes(df))
        z.writestr("process_map.md", markdown_bytes(process_map))
        z.writestr("document_risk_report.md", markdown_bytes(risk_report))
    return buffer.getvalue()

# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(page_title="AI Knowledge Inventory", page_icon="🧭", layout="wide")

st.markdown("""
<style>
.block-container { padding-top: 1.8rem; max-width: 1400px; }
.hero {
  padding: 1.25rem 1.35rem;
  border-radius: 22px;
  background: linear-gradient(135deg, rgba(37,99,235,0.12), rgba(14,165,233,0.08), rgba(255,255,255,0.9));
  border: 1px solid rgba(148, 163, 184, 0.28);
  margin-bottom: 1rem;
}
.hero-title { font-size: 2rem; font-weight: 800; letter-spacing: -0.03em; margin-bottom: 0.25rem; }
.hero-subtitle { color: #64748b; font-size: 1rem; }
.metric-card {
  padding: 1rem 1.1rem;
  border-radius: 18px;
  border: 1px solid rgba(148, 163, 184, 0.28);
  background: rgba(255, 255, 255, 0.82);
  box-shadow: 0 8px 24px rgba(15,23,42,0.05);
}
.metric-label { font-size: 0.85rem; color: #64748b; margin-bottom: 0.35rem; }
.metric-value { font-size: 1.9rem; font-weight: 800; }
.section-card {
  padding: 1rem 1.15rem;
  border-radius: 18px;
  border: 1px solid rgba(148, 163, 184, 0.28);
  background: rgba(255, 255, 255, 0.82);
  margin-bottom: 1rem;
}
.small-muted { color: #64748b; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <div class="hero-title">🧭 AI Knowledge Inventory</div>
  <div class="hero-subtitle">Upload messy enterprise documents, analyze knowledge readiness, and download a structured report.</div>
</div>
""", unsafe_allow_html=True)

default_config = get_default_ai_config()

with st.sidebar:
    st.header("Model")

    provider = st.selectbox(
        "Provider",
        ["No API", "OpenAI", "Gemini", "Claude", "Other"],
        index=0,
        help="Use No API for UI testing. Use Other for OpenAI-compatible providers."
    )

    api_key = ""
    base_url = ""
    model = ""
    mode = "heuristic"

    if provider == "No API":
        mode = "heuristic"
        st.info("Rule-based fallback. Useful for interface testing.")
    elif provider == "OpenAI":
        mode = "openai_compatible"
        api_key = st.text_input("API key", type="password", value=default_config.get("api_key", ""))
        model = st.text_input("Model", value="gpt-4o-mini")
        base_url = "https://api.openai.com/v1"
    elif provider == "Gemini":
        mode = "gemini"
        api_key = st.text_input("API key", type="password", value=default_config.get("api_key", ""))
        model = st.text_input("Model", value="gemini-2.5-flash")
        base_url = "https://generativelanguage.googleapis.com"
    elif provider == "Claude":
        mode = "anthropic"
        api_key = st.text_input("API key", type="password", value=default_config.get("api_key", ""))
        model = st.text_input("Model", value="claude-3-5-haiku-latest")
        base_url = "https://api.anthropic.com"
    elif provider == "Other":
        mode = "openai_compatible"
        api_key = st.text_input("API key", type="password", value=default_config.get("api_key", ""))
        base_url = st.text_input("Base URL", placeholder="https://api.deepseek.com/v1")
        model = st.text_input("Model", placeholder="deepseek-chat")
        st.caption("For DeepSeek, Qwen-compatible endpoints, Moonshot, local gateways, etc.")

    ai_config = {
        "mode": mode,
        "api_key": api_key,
        "base_url": base_url,
        "model": model,
    }

    st.markdown("---")
    st.header("Extraction")
    max_chars = st.slider("Text limit per document", 2000, 50000, 15000, step=1000)
    ocr_images = st.checkbox("Enable OCR for images/scanned PDFs", value=False)
    st.caption("OCR requires Tesseract. If unavailable, the app will continue and show extraction warnings.")

for key, default in {
    "inventory": None,
    "process_map": "",
    "risk_report": "",
    "raw_docs_count": 0,
    "extraction_warnings": [],
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

def run_analysis(extracted_docs):
    if not extracted_docs:
        st.error("No readable documents selected.")
        return

    st.session_state.raw_docs_count = len(extracted_docs)
    warnings = []
    results = []
    total = len(extracted_docs)

    progress = st.progress(0)
    status = st.empty()

    for idx, doc in enumerate(extracted_docs, start=1):
        status.info(f"Analyzing {idx}/{total}: {doc['document_name']}")
        if doc.get("extraction_warnings"):
            warnings.extend([f"{doc['document_name']}: {w}" for w in doc["extraction_warnings"]])

        metadata = {
            "document_name": doc["document_name"],
            "source_path": doc["source_path"],
            "file_type": doc["file_type"],
            "sheet_names": doc.get("sheet_names", []),
            "text_length": len(doc.get("text", "")),
            "extraction_warnings": doc.get("extraction_warnings", []),
        }

        results.append(
            analyze_document(
                doc.get("text", ""),
                metadata,
                max_chars=max_chars,
                ai_config=ai_config
            )
        )
        progress.progress(idx / total)

    status.info("Generating process map and risk report...")

    df = pd.DataFrame(results)
    records = df.to_dict(orient="records")
    process_map = generate_process_map(records, ai_config=ai_config)
    risk_report = generate_risk_report(records, ai_config=ai_config)

    st.session_state.inventory = df
    st.session_state.process_map = process_map
    st.session_state.risk_report = risk_report
    st.session_state.extraction_warnings = warnings

    status.success("Analysis complete.")
    progress.progress(1.0)

def metric_card(label, value):
    st.markdown(f"""
    <div class="metric-card">
      <div class="metric-label">{label}</div>
      <div class="metric-value">{value}</div>
    </div>
    """, unsafe_allow_html=True)

def render_dashboard(df: pd.DataFrame):
    if df is None or df.empty:
        return

    total = len(df)
    high = int((df["risk_level"] == "High").sum()) if "risk_level" in df else 0
    medium = int((df["risk_level"] == "Medium").sum()) if "risk_level" in df else 0
    review = int(df["needs_human_review"].astype(bool).sum()) if "needs_human_review" in df else 0
    kb = int((df["can_enter_kb"] == "Yes").sum()) if "can_enter_kb" in df else 0

    cols = st.columns(5)
    with cols[0]: metric_card("Documents", total)
    with cols[1]: metric_card("High risk", high)
    with cols[2]: metric_card("Medium risk", medium)
    with cols[3]: metric_card("Need review", review)
    with cols[4]: metric_card("KB candidates", kb)

    st.markdown("")
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Interactive dashboard")

    f1, f2, f3 = st.columns(3)
    risk_options = sorted(df["risk_level"].dropna().unique().tolist()) if "risk_level" in df else []
    dept_options = sorted(df["department"].dropna().unique().tolist()) if "department" in df else []
    kb_options = sorted(df["can_enter_kb"].dropna().unique().tolist()) if "can_enter_kb" in df else []

    risk_filter = f1.multiselect("Risk level", risk_options, default=risk_options)
    dept_filter = f2.multiselect("Department", dept_options, default=dept_options)
    kb_filter = f3.multiselect("KB readiness", kb_options, default=kb_options)

    filtered = df.copy()
    if risk_filter and "risk_level" in filtered:
        filtered = filtered[filtered["risk_level"].isin(risk_filter)]
    if dept_filter and "department" in filtered:
        filtered = filtered[filtered["department"].isin(dept_filter)]
    if kb_filter and "can_enter_kb" in filtered:
        filtered = filtered[filtered["can_enter_kb"].isin(kb_filter)]

    chart1, chart2 = st.columns(2)

    with chart1:
        if not filtered.empty and "risk_level" in filtered:
            order = ["High", "Medium", "Low"]
            counts = filtered["risk_level"].value_counts().reindex(order).dropna().reset_index()
            counts.columns = ["risk_level", "count"]
            fig = px.bar(
                counts,
                x="risk_level",
                y="count",
                title="Risk Distribution",
                text="count",
                color="risk_level",
                color_discrete_map={"High": "#ef4444", "Medium": "#f59e0b", "Low": "#22c55e"},
            )
            fig.update_layout(height=350, margin=dict(l=10, r=10, t=60, b=10))
            st.plotly_chart(fig, use_container_width=True)

    with chart2:
        if not filtered.empty and "department" in filtered:
            counts = filtered["department"].value_counts().head(10).reset_index()
            counts.columns = ["department", "count"]
            fig = px.bar(
                counts,
                x="count",
                y="department",
                orientation="h",
                title="Documents by Department",
                text="count",
                color="count",
                color_continuous_scale="Blues",
            )
            fig.update_layout(height=350, margin=dict(l=10, r=10, t=60, b=10), yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

    chart3, chart4 = st.columns(2)

    with chart3:
        if not filtered.empty and "document_type" in filtered:
            counts = filtered["document_type"].value_counts().head(8).reset_index()
            counts.columns = ["document_type", "count"]
            fig = px.pie(counts, names="document_type", values="count", title="Document Type Mix", hole=0.55)
            fig.update_layout(height=380, margin=dict(l=10, r=10, t=60, b=10))
            st.plotly_chart(fig, use_container_width=True)

    with chart4:
        if not filtered.empty and "can_enter_kb" in filtered:
            counts = filtered["can_enter_kb"].value_counts().reset_index()
            counts.columns = ["can_enter_kb", "count"]
            fig = px.pie(
                counts,
                names="can_enter_kb",
                values="count",
                title="Knowledge Base Readiness",
                hole=0.55,
                color="can_enter_kb",
                color_discrete_map={"Yes": "#22c55e", "Needs Review": "#f59e0b", "No": "#ef4444"},
            )
            fig.update_layout(height=380, margin=dict(l=10, r=10, t=60, b=10))
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Priority Review Items")
    review_df = filtered.copy()
    if "needs_human_review" in review_df:
        review_df = review_df[review_df["needs_human_review"].astype(bool)]
    if "risk_level" in review_df:
        review_df["_risk_order"] = review_df["risk_level"].map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        review_df = review_df.sort_values("_risk_order").drop(columns=["_risk_order"])

    cols = ["document_name", "department", "business_process", "risk_level", "risk_reason", "recommended_action"]
    cols = [c for c in cols if c in review_df.columns]
    if not review_df.empty:
        st.dataframe(review_df[cols], use_container_width=True, hide_index=True)
    else:
        st.success("No priority review items in the current filter.")
    st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.extraction_warnings:
        with st.expander("Extraction warnings"):
            for warning in st.session_state.extraction_warnings:
                st.write(f"- {warning}")

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("1. Add documents")
st.markdown('<div class="small-muted">Recommended for folder upload: compress the folder into ZIP and upload it here.</div>', unsafe_allow_html=True)

tab_zip, tab_files = st.tabs(["Upload folder ZIP", "Upload files"])

zip_file = None
uploaded_files = None

with tab_zip:
    zip_file = st.file_uploader("Upload a ZIP file", type=["zip"], accept_multiple_files=False)
    if zip_file:
        try:
            preview_docs = extract_docs_from_zip_upload(zip_file, ocr_images=ocr_images)
            st.success(f"{len(preview_docs)} supported document(s) found in ZIP.")
        except Exception as exc:
            st.error(f"Could not read ZIP: {exc}")

with tab_files:
    uploaded_files = st.file_uploader(
        "Upload individual files",
        type=[ext.replace(".", "") for ext in SUPPORTED_EXTENSIONS],
        accept_multiple_files=True
    )
    if uploaded_files:
        st.success(f"{len(uploaded_files)} file(s) uploaded.")

st.markdown("</div>", unsafe_allow_html=True)

selected_count = 0
if zip_file:
    try:
        selected_count += len(extract_docs_from_zip_upload(zip_file, ocr_images=ocr_images))
    except Exception:
        pass
if uploaded_files:
    selected_count += len(uploaded_files)

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("2. Run analysis")
if selected_count:
    st.info(f"{selected_count} document(s) selected.")
else:
    st.warning("No documents selected yet.")

if st.button("Run Knowledge Inventory Analysis", type="primary", use_container_width=True):
    extracted = []
    if zip_file:
        extracted.extend(extract_docs_from_zip_upload(zip_file, ocr_images=ocr_images))
    if uploaded_files:
        extracted.extend([extract_text_from_upload(f, ocr_images=ocr_images) for f in uploaded_files])

    seen = set()
    unique_docs = []
    for doc in extracted:
        key = (doc.get("document_name"), doc.get("source_path"))
        if key not in seen:
            seen.add(key)
            unique_docs.append(doc)

    run_analysis(unique_docs)
st.markdown("</div>", unsafe_allow_html=True)

if st.session_state.inventory is not None:
    st.markdown("---")
    render_dashboard(st.session_state.inventory)

    st.markdown("---")
    tab1, tab2, tab3, tab4 = st.tabs(["Inventory table", "Process map", "Risk report", "Download reports"])

    with tab1:
        display_cols = [
            "document_name", "file_type", "document_type", "department", "business_process",
            "knowledge_category", "risk_level", "can_enter_kb",
            "needs_human_review", "summary", "recommended_action"
        ]
        display_cols = [c for c in display_cols if c in st.session_state.inventory.columns]
        st.dataframe(st.session_state.inventory[display_cols], use_container_width=True, hide_index=True)

    with tab2:
        st.markdown(st.session_state.process_map)

    with tab3:
        st.markdown(st.session_state.risk_report)

    with tab4:
        df = st.session_state.inventory
        process_map = st.session_state.process_map
        risk_report = st.session_state.risk_report

        d1, d2, d3, d4 = st.columns(4)
        with d1:
            st.download_button(
                "Download Excel",
                data=inventory_excel_bytes(df),
                file_name="knowledge_inventory.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "Download Process Map",
                data=markdown_bytes(process_map),
                file_name="process_map.md",
                mime="text/markdown",
                use_container_width=True,
            )
        with d3:
            st.download_button(
                "Download Risk Report",
                data=markdown_bytes(risk_report),
                file_name="document_risk_report.md",
                mime="text/markdown",
                use_container_width=True,
            )
        with d4:
            st.download_button(
                "Download All",
                data=report_zip_bytes(df, process_map, risk_report),
                file_name="knowledge_inventory_report_bundle.zip",
                mime="application/zip",
                use_container_width=True,
            )
