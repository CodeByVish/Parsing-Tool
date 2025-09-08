# Parsing-Tool-
# Gradiant KB → GPT Conversion Pipeline

*A single-file CLI (`kb_tool.py`) that turns messy enterprise documents into accurate, auditable inputs for a Custom GPT. Built for Windows (no admin), works on macOS/Linux too.*

---

## ✨ What it does

* **Converts**: `PDF`, `PPTX`, `DOCX`, `XLS/XLSX`, `PNG/JPG` → clean `.txt`
* **Preserves tables** from slides/docs as TSV (copy-paste into Excel/Pandas)
* **Captures figures**: OCR text + a **PromptSuggestion** to regenerate a clean PNG (flow diagram / chart / illustration) inside ChatGPT
* **Anchors everything** with file path + page/slide markers for auditability
* **Bundles** many `.txt` into big files for better retrieval in Custom GPTs
* **No Mermaid noise** (v2 removed it after evaluation)
* **Runs offline** (no web calls). Optional Tesseract OCR for images/scans.

---

## 🗂 Repo layout

```
.
├─ kb_tool.py                  # the CLI (convert, bundle, split, doctor)
├─ requirements.txt            # Python deps
├─ README.md                   # this file
└─ examples/
   ├─ flow_from_gpt.csv        # sample CSV for draw.io testing (optional)
   └─ drawio_csv_config.txt    # sample config text for CSV import (optional)
```

> If you don’t see `examples/`, copy the snippets below from this README.

---

## 🚀 Quick start (Windows, **no admin**)

1. **Create a workspace & venv**

```bat
C:
mkdir C:\kb
cd C:\kb
python -m venv .venv
C:\kb\.venv\Scripts\activate.bat
```

2. **Install dependencies**

```bat
pip install -r requirements.txt
```

3. **Doctor check (set your Tesseract path)**

```bat
python kb_tool.py doctor --tesseract "C:\Users\<YOU>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
```

> If you don’t have Tesseract, install it first (no admin installer works as well).

4. **Convert your source folder**

```bat
python kb_tool.py convert ^
  --input  C:\kb\raw ^
  --output C:\kb\output ^
  --tesseract "C:\Users\<YOU>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
```

5. **Bundle outputs for GPT upload**

```bat
python kb_tool.py bundle --output C:\kb\output
```

You’ll get one or more files like:

```
C:\kb\output\bundle__Engineering-application-data.txt
```

Upload those bundle files to your Custom GPT.

---

## 🛠 CLI reference

```bat
python kb_tool.py convert --input <folder> --output <folder> [--tesseract <path>] [--ocr-lang eng]
python kb_tool.py bundle  --output <folder>
python kb_tool.py split-text --file <big.txt> --max-mb 100
python kb_tool.py doctor --tesseract <path>
```

**Useful flags**

* `--ocr-lang` — e.g., `eng`, or multi-lang `"eng+ara"`
* By design v2 writes **one .txt per source file** (no token chunking). Use `split-text` later only if a bundle exceeds an upload limit.

---

## 📦 Output format & guarantees

* **Traceable Anchors**

  * `[SOURCE] C:\path\to\file.pptx`
  * `=== Slide N ===` / `=== PDF Page N ===`
* **Tables preserved**

  * Blocks start with `[Table N]`
  * Rows are **tab-separated** (DataFrame-ready), units retained
* **Figures captured**

  * `[Picture on Slide N | TYPE]`
  * `OCR:` short text recovered from the image
  * `PromptSuggestion:` how to recreate a clean PNG (chart/flow/illustration) in ChatGPT
* **No Mermaid artifacts**

  * Visuals are regenerated from text evidence inside ChatGPT
* **Audit & compliance**

  * Every GPT answer must end with **Sources** (file → slide/page → block)

---

## 🧪 External testing (for managers & reviewers)

### Option A — draw\.io (diagrams.net) **Mermaid** importer (fastest)

1. In app.diagrams.net: **Arrange → Insert → Mermaid…**
2. Paste a flowchart (from your GPT spec). Example:

   ```mermaid
   flowchart LR
     S1["Combine LSR and HFW"] -->|Mixing| S2["Send to RO membrane"]
     S2 -->|Feed| S3["Produce clean permeate"]
     S3 -->|Permeate| S4["Generate small brine"]
     S4 -->|Reject| S5["Brine from service cycle"]
     S5 -->|Brine| S6["Brine from flushing cycle"]
   ```
3. Export **PNG** and attach the GPT **Sources** for provenance.

### Option B — draw\.io **CSV** import (structured)

* **File → Import From → CSV…**
* **Top box:** paste your CSV (from the GPT)
* **Bottom box:** paste this configuration

  ```
  # label: %Process Step Description%
  # style: rounded=0;whiteSpace=wrap;html=1;fontSize=14;strokeColor=#34495E;fillColor=#FFFFFF
  # parent: %Function%
  # parentstyle: swimlane;whiteSpace=wrap;html=1;childLayout=stackLayout;horizontal=1;horizontalStack=0;resizeParent=1;collapsible=1;
  # connect: {"from": "Process Step ID", "to": "Next Step ID", "label": "Connector Label",
  #           "style": "endArrow=block;endFill=1;strokeColor=#34495E;fontSize=12;"}
  # layout: auto
  # nodespacing: 60
  # levelspacing: 100
  # edgespacing: 40
  # ignore: Phase
  ```
* **CSV example**

  ```
  Process Step ID,Process Step Description,Next Step ID,Connector Label,Function,Phase
  S1,Combine LSR and HFW,S2,Mixing,Equalization,Pre-treatment
  S2,Send to RO membrane,S3,Feed,Separation,Main treatment
  S3,Produce clean permeate,S4,Permeate,Water recovery,Main treatment
  S4,Generate small brine,S5,Reject,Concentration,Main treatment
  S5,Brine from service cycle,S6,Brine,Discharge,ZLD interface
  S6,Brine from flushing cycle,,Brine,Discharge,ZLD interface
  ```

**Pass criteria**

* Nodes & edges match the process order (branches okay)
* Connector labels present where provided
* Keep & share the CSV + PNG + **Sources**

---

## 🤖 Suggested Custom GPT instructions (drop-in)

**Task**
Answer questions using only the uploaded bundle files. Reconstruct tables and visuals as needed. **Always include sources.**

**Persona**
A careful, audit-oriented technical assistant for Gradiant. Prioritizes data fidelity and provenance.

**Rules**

1. Use only the uploaded bundles as your knowledge base.
2. If asked for tables, return a clean table (or CSV/TSV) from `[Table N]` blocks.
3. If asked for a chart, first return CSV data, then render a simple PNG chart (axes titled with units).
4. If asked for a process flow, first return a short JSON spec `{nodes:[], edges:[]}`, then render a left-to-right PNG diagram.
5. **Sources (mandatory)**: end every answer with

   ```
   Sources:
   <FILE path> → Slide N → [Table N] / [Picture on Slide N | TYPE]
   ```
6. If evidence is missing, say **“Sources: none — not found in uploaded bundle”** and ask ≤2 precise follow-ups.
7. Do not invent values or equipment. Prefer **unknown** to guessing.

**Conversation starters**

* “Summarize the process on Slide 4 and draw a simple flow.”
* “Extract parameters and units from Table 2 into CSV.”
* “Build a chart of IPA concentration vs. RI from the data on Slide 7.”

---

## 📈 Quality rubric (internal)

| Dimension                     | Target                           |
| ----------------------------- | -------------------------------- |
| **Sources present & correct** | 100% of answers                  |
| **Tables (values & units)**   | ≥ 99% exact match                |
| **Charts (CSV match)**        | ≥ 99% values; correct axis units |
| **Flows (recall)**            | ≥ 85% nodes, ≥ 80% edges         |
| **Repeatability**             | Same JSON/CSV across two runs    |

---

## 🧩 Requirements

`requirements.txt`

```
pymupdf>=1.23.0
python-pptx>=0.6.21
python-docx>=1.1.0
pandas>=2.0.0
openpyxl>=3.1.2
Pillow>=10.0.0
pytesseract>=0.3.10
tiktoken>=0.7.0
tqdm>=4.66.0
```

---

## 🧯 Troubleshooting

* **Execution policy blocks venv activation (PowerShell):**
  Use CMD: `C:\kb\.venv\Scripts\activate.bat` (no policy change needed).

* **Tesseract not found / wrong path:**
  Run `python kb_tool.py doctor --tesseract "<path>\tesseract.exe"`.

* **`~$filename.pptx` extraction error:**
  That’s a PowerPoint lock file. Close the PPT and delete the `~$` temp file, or ignore.

* **`MSO_SHAPE_TYPE.CONNECTOR` errors:**
  v2 no longer relies on connectors; figures are handled via OCR + PromptSuggestion.

* **Slow conversion:**
  Large scans/figures require OCR. You can temporarily remove images or run per-subfolder to parallelize across machines.

* **Huge bundles:**
  Use `split-text` to create smaller parts (e.g., 100 MB).

---

## 🔒 Privacy

* The converter runs locally and never calls external services.
* Only upload bundles to systems approved by your org.

---

## 📝 License

Choose a license for your repo (e.g., MIT). Example:

```
MIT License
Copyright (c) 2025 Gradiant
```

---

## 🤝 Contributing

Issues and PRs welcome. Please:

* Describe the source file and a minimal reproduction
* Attach a redacted page/slide if possible
* Include your `doctor` output and OS/Python versions

---

## 🧭 Changelog

* **v2**: Removed Mermaid; added figure typing + OCR PromptSuggestion; one `.txt` per file; improved anchors; Windows CMD flow; external testing guide (draw\.io).

---

### Appendix — sample files

**`examples/flow_from_gpt.csv`**

```
Process Step ID,Process Step Description,Next Step ID,Connector Label,Function,Phase
S1,Combine LSR and HFW,S2,Mixing,Equalization,Pre-treatment
S2,Send to RO membrane,S3,Feed,Separation,Main treatment
S3,Produce clean permeate,S4,Permeate,Water recovery,Main treatment
S4,Generate small brine,S5,Reject,Concentration,Main treatment
S5,Brine from service cycle,S6,Brine,Discharge,ZLD interface
S6,Brine from flushing cycle,,Brine,Discharge,ZLD interface
```

**`examples/drawio_csv_config.txt`**

```
# label: %Process Step Description%
# style: rounded=0;whiteSpace=wrap;html=1;fontSize=14;strokeColor=#34495E;fillColor=#FFFFFF
# parent: %Function%
# parentstyle: swimlane;whiteSpace=wrap;html=1;childLayout=stackLayout;horizontal=1;horizontalStack=0;resizeParent=1;collapsible=1;
# connect: {"from": "Process Step ID", "to": "Next Step ID", "label": "Connector Label",
#           "style": "endArrow=block;endFill=1;strokeColor=#34495E;fontSize=12;"}
# layout: auto
# nodespacing: 60
# levelspacing: 100
# edgespacing: 40
# ignore: Phase
```

---

If you want, I can also generate a minimal **sample repo** structure you can push directly to GitHub (with `requirements.txt`, the examples folder, and this README).
