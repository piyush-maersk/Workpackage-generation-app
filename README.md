# Workpackage Generation App

> **Prototype v0.1** — Automation & Operational Technology  
> RAG-powered automated workpackage document generation

---

## Overview

This prototype automates the creation of workpackage documents for warehouse
hardware installation and configuration projects.  
It uses a **Retrieval-Augmented Generation (RAG)** pipeline to analyse input
documents (FbM Tech Estimation Excel sheets and scope-of-work text) and
generate populated DOCX workpackages ready for review.

---

## Supported Templates

| Template | Status | Category |
|---|---|---|
| IT Device | ✅ Available | Hardware |
| OT Device | ✅ Available | Hardware |
| MDF | ✅ Available | Network |
| OT Automation Machine | ✅ Available | Automation |
| Perimeter | 🔜 Yet to Arrive | Network |
| IT Net | 🔜 Yet to Arrive | Network |
| OT Net | 🔜 Yet to Arrive | Network |
| Automation | 🔜 Yet to Arrive | Automation |

---

## Folder Structure

```
Workpackage-generation-app/
├── app.py                              # Streamlit UI
├── src/
│   ├── __init__.py
│   ├── parser.py                       # Scope-of-work text + FbM Excel parser
│   ├── device_classifier.py            # Keyword-based device categorisation
│   ├── rag_engine.py                   # LangChain + OpenAI + ChromaDB RAG pipeline
│   └── template_filler.py             # DOCX template population
├── requirements.txt                    # Python dependencies
├── .env.example                        # Environment variable template
├── README.md                           # This file
│
│   ── Workpackage Templates (existing) ──
├── IT Hardware - Workpackage Template.docx
├── OT Hardware - Workpackage Template.docx
├── MDF_Workpackage.docx
└── OT_Automationmachine_Workpackage.docx
```

---

## Setup

### 1. Prerequisites

- Python 3.10 or higher
- An [OpenAI API key](https://platform.openai.com/api-keys) (GPT-4o or GPT-4 recommended)

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure environment

```bash
cp .env.example .env
# Edit .env and set your OPENAI_API_KEY
```

Alternatively, enter the API key directly in the app's sidebar.

### 4. Run the app

```bash
streamlit run app.py
```

Open **http://localhost:8501** in your browser.

---

## Usage

1. **Select Template** — Choose one of the 4 available workpackage types.
2. **Fill Project Information** — Enter any known project details (project name,
   FBM ID, site address, etc.).  The AI will attempt to auto-extract whatever
   is left blank from your input documents.
3. **Provide Input Data** (at least one required):
   - *Scope of Work tab* — Paste text or upload a `.txt` file.
   - *Cost Estimation tab* — Upload the FbM Tech Estimation `.xlsx` file.
4. **Generate** — Click the button and wait ~15–30 seconds.
5. **Download** — Download the generated `.docx` workpackage.

---

## Architecture

```
Input Data                        RAG Pipeline                    Output
─────────────────────────────────────────────────────────────────────────────
Scope of Work (text / .txt)  ─┐
                               ├─► InputParser ──► DeviceClassifier ─────────┐
FbM Estimation (.xlsx)       ─┘                                               │
                                                                              ▼
DOCX Templates ──► OpenAI Embeddings ──► ChromaDB ──► Retriever ──► LLM (GPT-4)
                   (indexed on startup)                                       │
                                                                              ▼
                                                               TemplateFiller (.docx)
                                                                              │
                                                                              ▼
                                                               Generated Workpackage
```

### Key Components

| Component | File | Purpose |
|---|---|---|
| Streamlit UI | `app.py` | User interface for inputs and download |
| Input Parser | `src/parser.py` | Extract device list from text and Excel |
| Device Classifier | `src/device_classifier.py` | Categorise devices as IT/OT/Network/MDF/Automation |
| RAG Engine | `src/rag_engine.py` | Index templates, extract project info, generate section content |
| Template Filler | `src/template_filler.py` | Populate DOCX with project data and generated content |

---

## Device Categories

| Category | Example Devices |
|---|---|
| IT | Laptops, Desktops (admin/office), Monitors, Docking Stations, MFP Printers, Phones |
| OT | RF Scanners, Label Printers, Mobile Printers, Rugged Tablets, Battery Chargers |
| Network | Access Switches, Access Points, Firewalls, Patch Panels, UPS (1400 VA) |
| MDF | Core Switches, Server Racks, PDUs, UPS (2200 VA+) |
| Automation | RFID Tunnels, Cubiscan, PANDA Systems, Carton Erectors |
| Software | SOTI Licences, Bartender, WMS/DMS Subscriptions |

---

## Notes

- This is a **prototype** focused on extraction accuracy and template-filling
  correctness.  It is not intended for production use without further review.
- The ChromaDB vector store is built in-memory on every run (no disk
  persistence required for the prototype).
- Cost and pricing information from estimation sheets is intentionally ignored;
  only device names and quantities are extracted.
- GPT-4o / GPT-4 strongly recommended for best extraction and generation quality.
  GPT-3.5-Turbo may produce lower-quality output.
