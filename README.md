# AI Knowledge Intake Agent

### An agentic knowledge-readiness layer before enterprise AI systems

Most enterprise AI projects do not fail because the model is weak.
They fail because the knowledge going into the model is messy, outdated, unvalidated, or impossible to trace.

**AI Knowledge Intake Agent** is a prototype that explores the missing layer between enterprise documents and AI assistants.

Instead of starting with a chatbot, this project starts one step earlier:

> Can the company’s internal knowledge be trusted, corrected, and prepared before it enters a knowledge base, RAG system, or enterprise AI assistant?

---

## Live Demo

**Streamlit App:** https://knowledge-inventory-vcoybrpmsusfcypohrb9vc.streamlit.app/

**Contact:** kexin.lyu@edhec.com

---

## Why this project exists

Many companies are trying to build AI assistants on top of internal documents.

But in real enterprise environments, documents are often scattered across Word files, Excel sheets, PDFs, SOPs, meeting notes, manuals, and local folders.

Common problems include:

* No clear document owner
* Missing or outdated version information
* Conflicting business rules across files
* Meeting notes mixed with official process documents
* Unclear source-of-truth status
* Incomplete extraction from complex files
* No audit trail for AI-generated corrections

If these documents are directly ingested into a RAG system, the AI may produce answers that sound confident but are based on unverified knowledge.

This project asks a more fundamental question:

> Before enterprise AI answers, who checks whether the knowledge itself is ready?

---

## What it does

AI Knowledge Intake Agent helps turn messy business documents into AI-ready knowledge assets through a guided, human-in-the-loop workflow.

The system can:

* Upload multiple documents or a ZIP folder
* Scan and classify business documents
* Build a knowledge inventory
* Diagnose document-level risks
* Identify missing owners, unclear versions, weak source-of-truth signals, and other readiness issues
* Recommend the next best workflow step through an Agent Planner
* Generate guided pre-correction questions
* Offer recommended answer options while allowing custom user input
* Apply corrections at batch, segment, selected-document, or single-document level
* Generate corrected document drafts
* Produce a correction ledger for traceability
* Export corrected reports and knowledge-ready outputs

---

## The core idea

This is not another chatbot demo.

It is a prototype for an **enterprise knowledge intake layer**.

```text
Messy enterprise documents
        ↓
Knowledge readiness scan
        ↓
Risk diagnosis
        ↓
Agent planner
        ↓
Scope-aware pre-correction
        ↓
Human-confirmed corrections
        ↓
Corrected documents + ledger
        ↓
AI-ready knowledge assets
```

The goal is not to let AI replace human judgment.

The goal is to let AI surface uncertainty, ask better questions, reduce manual review effort, and preserve evidence for every correction.

---

## Key product concept: Scope-aware Pre-correction

In enterprise workflows, users cannot review every file one by one.

This project introduces **scope-aware pre-correction**, where the system adapts the correction process based on the user’s selected scope.

Users can choose:

| Correction scope             | Use case                                                   |
| ---------------------------- | ---------------------------------------------------------- |
| Batch-level correction       | Fix shared issues across many documents                    |
| Segment-level correction     | Refine a department, process, document type, or risk group |
| Selected-document correction | Correct a chosen set of documents                          |
| Document-level correction    | Precisely correct one important file                       |

Each clarification question includes:

* The issue found
* Why it matters
* AI’s assumption
* Recommended answer option
* Alternative options
* A custom input field
* Apply-scope control
* Correction ledger output

This allows users to correct many documents with a few high-value decisions, while still preserving a path for precise single-file correction.

---

## Agent Planner

The latest version includes an **Agent Planner** that recommends the next best action after the initial diagnosis.

For example, the system may recommend:

* Start with batch-level correction if many documents share the same issue
* Refine a specific department or process if risks are concentrated in one segment
* Review a single critical SOP if it cannot be safely corrected by batch rules
* Export corrected documents if the knowledge is already ready enough

The agent does not publish or overwrite source documents automatically.

It recommends the next workflow step, explains why, and waits for user confirmation.

---

## Outputs

After analysis and pre-correction, the app can generate:

* `knowledge_inventory_corrected.xlsx`
* `document_risk_report_corrected.md`
* `process_map_corrected.md`
* `correction_ledger.xlsx`
* `corrected_documents.zip`
* `knowledge_inventory_corrected_bundle.zip`

Original source files are preserved as evidence.
Corrected drafts are generated as separate outputs.

---

## Why this matters for enterprise AI

Enterprise AI adoption is not just a model problem.
It is a knowledge operations problem.

A reliable enterprise AI system needs:

* Clean knowledge intake
* Human validation
* Source traceability
* Version awareness
* Correction history
* Controlled publishing
* Governance before RAG ingestion

This project explores how AI can support that workflow.

The principle is simple:

> AI prepares the work.
> Humans confirm the business truth.
> The system keeps the evidence.

---

## Current prototype

This version is built as a Streamlit demo for fast experimentation and portfolio presentation.

The same logic can be adapted into an API-first architecture for enterprise integration.

Potential enterprise deployment patterns include:

* FastGPT workflow node
* SharePoint document intake service
* Internal portal integration
* API-based knowledge readiness scanner
* Batch document review pipeline
* RAG preparation middleware

In a production environment, the Streamlit interface would become a control room or demo shell, while the core capability would be exposed as APIs.

---

## Tech stack

* Python
* Streamlit
* Pandas
* OpenPyXL
* python-docx
* PDF / Office document parsing
* LLM-assisted classification and reasoning
* Human-in-the-loop correction workflow
* Exportable reports and document drafts

---

## What I am exploring next

This prototype is evolving toward an **agentic knowledge intake service** for enterprise AI.

Next areas of exploration include:

* API-first backend with FastAPI
* SharePoint / enterprise document source integration
* Permission-aware document processing
* Better source evidence binding
* Conflict detection across document versions
* Knowledge object extraction
* RAG-ready Markdown / JSON export
* Approval workflow before knowledge publishing
* Enterprise audit and governance layer

---

## Who this is for

This project may be relevant if your team is working on:

* Enterprise AI adoption
* RAG implementation
* Internal knowledge base modernization
* Business process documentation
* AI assistant reliability
* Document governance
* AI transformation workflows
* Human-in-the-loop AI systems

If you are exploring how to move from messy enterprise documents to reliable AI-ready knowledge, I would be happy to connect.

---

## One-line summary

**AI Knowledge Intake Agent is an agentic workflow prototype that scans messy enterprise documents, diagnoses knowledge readiness, asks targeted clarification questions, applies human-confirmed corrections, and generates traceable AI-ready knowledge assets before RAG or enterprise AI ingestion.**
