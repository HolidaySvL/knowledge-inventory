# AI Knowledge Inventory & Collaborative Workflow 

# URL：https://knowledge-inventory-vcoybrpmsusfcypohrb9vc.streamlit.app/

## Project Overview

**AI Knowledge Inventory** is a business-oriented enterprise AI workflow designed to help organizations prepare messy internal documents before they enter a knowledge base, RAG system, or enterprise AI assistant.

Instead of building a chatbot directly on top of unvalidated documents, this project focuses on the earlier and often overlooked layer of enterprise AI adoption: **knowledge preparation, document quality review, validation, and knowledge governance**.

The project includes a deployed Streamlit web app that allows users to upload business documents or a ZIP folder, analyze document readiness, identify risks, generate an interactive dashboard, and download structured knowledge inventory reports.

---

## Problem

Many companies want to build AI assistants, internal chatbots, or RAG-based knowledge systems. However, enterprise documents are often messy, inconsistent, and difficult to trust.

Common problems include:

* Documents are scattered across Word files, Excel sheets, PDFs, meeting notes, manuals, and local folders.
* Many documents have unclear owners, versions, or approval status.
* Some documents are outdated, duplicated, incomplete, or inconsistent.
* Business rules are hidden inside long documents rather than structured as reusable knowledge.
* Directly putting messy documents into a RAG system can lead to inaccurate, untraceable, or conflicting AI answers.

The core insight behind this project is:

> Enterprise AI does not only need better chatbots. It needs a reliable knowledge preparation layer before information enters the AI system.

---

## Solution

This project proposes a three-stage enterprise AI knowledge workflow:

### Stage 1: Review Existing Documents

The system first helps review and inventory existing business materials, such as SOPs, manuals, FAQs, Excel records, meeting notes, PDFs, and training documents.

The goal is to understand what knowledge assets already exist and whether they are suitable for AI use.

Key outputs include:

* Document inventory
* Document type classification
* Department and process mapping
* Risk level assessment
* Document readiness status
* Open questions for human review

### Stage 2: Process Qualified Documents

Only documents that pass a quality gate should move into deeper processing.

For qualified documents, the system can support:

* Structured document chunking
* Field and metadata extraction
* Naming and tagging standardization
* FAQ draft generation
* Knowledge card preparation
* Process step extraction
* Table, attachment, and image relationship identification

This stage is designed to transform business documents into reusable and traceable knowledge assets.

### Stage 3: Knowledge Ingestion

After review and processing, cleaned knowledge can be prepared for:

* Knowledge base ingestion
* RAG systems
* Enterprise AI assistants
* Internal search
* Training and onboarding
* Business process reuse

The final goal is not only to answer questions, but to create a sustainable knowledge management workflow.

---

## Collaborative Workflow Concept

Beyond one-time document cleanup, this project also introduces a collaborative workflow for continuous knowledge generation.

The workflow follows a top-down and bottom-up loop:

### Top Layer: Strategy & Goals

Leadership defines direction, goals, knowledge standards, and business priorities.

### Middle Layer: Detailing & Review

Middle-level teams translate strategic goals into detailed rules, process definitions, validation criteria, and review checkpoints.

### Frontline Layer: Execution & Feedback

Frontline teams produce documents, record execution results, capture operational issues, and provide feedback.

The key idea is that document-processing capability should be embedded into the daily business workflow. As teams continue to work, they continuously generate new documents, records, FAQs, cases, and methods. These outputs can then be reviewed, validated, and accumulated into the enterprise knowledge base.

---

## Validation Framework

A major design principle of this project is that AI output should not be treated as final knowledge by default.

AI can read, organize, classify, and draft. However, critical judgment should remain with humans.

The validation framework includes:

### Rule Checks

The system checks whether documents contain key information such as owner, version, process scope, document type, and risk indicators.

### Human Review

Important or high-risk outputs are flagged for human confirmation instead of being automatically published.

### Source Traceability

Knowledge entries should remain connected to their original source documents, so users can verify where the information came from.

### Version Management

Documents with unclear, outdated, or conflicting versions are marked as requiring review.

### AI Self-Check

Before and after generation, the AI workflow should check for:

* Missing inputs
* Incomplete metadata
* Logical inconsistencies
* Conflicting information
* Unsupported conclusions
* High-risk knowledge claims

The principle is:

> AI prepares the work. Humans make the judgment. The system keeps the evidence.

---

## Product Features

The current MVP includes:

* Public Streamlit web app deployment
* Public GitHub repository
* Multi-file upload
* ZIP folder upload
* Support for multiple document formats
* Document text extraction
* AI-assisted document classification
* Risk level assessment
* Knowledge-base readiness assessment
* Interactive dashboard
* Downloadable Excel report
* Downloadable process map
* Downloadable risk report
* Downloadable report bundle

The dashboard helps users quickly understand:

* How many documents were analyzed
* Which documents are high-risk
* Which departments and processes are represented
* Which documents need human review
* Which documents are potential knowledge-base candidates

---

## Tech Stack

* Python
* Streamlit
* Pandas
* Plotly
* OpenPyXL
* PyMuPDF
* Python-docx
* Python-pptx
* OCR support through Tesseract
* LibreOffice-based legacy document conversion support
* LLM API integration with multiple provider options

---

## My Role

I designed and built this project end to end, including:

* Identifying the enterprise AI adoption problem
* Designing the knowledge preparation workflow
* Defining the three-stage document-to-knowledge process
* Designing the validation framework
* Building the Streamlit web application
* Implementing document upload and parsing logic
* Creating the interactive dashboard
* Designing downloadable reporting outputs
* Deploying the app publicly
* Preparing the project for portfolio and stakeholder demonstration

---

## Business Value

This project demonstrates how enterprise AI can be introduced more safely and practically by focusing on knowledge readiness before chatbot deployment.

Potential business value includes:

* Reducing the risk of inaccurate AI answers caused by poor source documents
* Helping companies understand their existing knowledge assets
* Identifying document quality issues before AI ingestion
* Supporting knowledge governance and validation
* Reducing manual document review workload
* Creating a bridge between business teams and technical AI implementation
* Supporting future RAG, enterprise assistant, and workflow automation initiatives

---

## Limitations

This MVP is a prototype and has several limitations:

* AI-generated analysis still requires human review.
* Document extraction quality depends on file format and document structure.
* OCR and legacy Office conversion depend on deployment environment.
* The system does not yet include enterprise permission control.
* It does not yet directly sync with SharePoint, Notion, FastGPT, or internal knowledge bases.
* It currently focuses on knowledge inventory rather than full automated knowledge publishing.

These limitations are intentional because the MVP focuses on validating the business workflow first.

---

## Roadmap

Future development could include:

* Human-in-the-loop review queue
* Source citation and evidence binding
* Conflict detection across documents
* Version comparison
* Knowledge object extraction
* FAQ and process card generation
* Direct export to RAG-ready Markdown
* Integration with SharePoint, Notion, FastGPT, or enterprise knowledge platforms
* Permission-aware document processing
* Workflow-based validation approval
* Department-level knowledge dashboards

---

## Core Principle

This project is built around one core principle:

> AI should not replace human judgment in enterprise knowledge work.
> AI should help humans read faster, organize better, validate more systematically, and turn messy information into reusable knowledge assets.
