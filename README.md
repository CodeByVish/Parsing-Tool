# Parsing-Tool-
# Gradiant FlowSmith — Doc→GPT Conversion Pipeline
This repo contains `kb_tool.py` and assets that turn PDFs/PPTX/DOCX/XLS(X)/images into **audit-ready bundles** consumed by the custom GPT **Gradiant FlowSmith**. Tables are preserved (TSV), figures get OCR + PromptSuggestion, and every snippet carries a source anchor (file → slide/page).

## About the GPT (Gradiant FlowSmith)
A cautious, evidence-first assistant that answers **only** from the uploaded bundles and **always ends with Sources**.

### Conversation starters
- “Summarize the process on Slide 4 and draw a simple flow.”
- “Extract Table 2 into CSV with units.”
- “Plot IPA concentration vs RI from the data on Slide 7.”
- “List assumptions or missing info for the pretreatment section.”
- “Compare the two configurations mentioned across Slides 10–12—cite each.”


---

## 🗂 Repo layout

```
.
├─ kb_tool.py                  # the CLI (convert, bundle, split, doctor)
├─ requirements.txt            # Python deps
├─ README.md                   # this file

```


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

## 🤖 Custom GPT instructions 

**Task**
Build answers, tables, charts, and simple diagrams strictly from the uploaded engineering knowledge base (the bundle__*.txt files). When useful or requested, reconstruct visuals (flows/charts) as clean PNGs using the table data and OCR blocks embedded in those files. **Always include sources.**

**Persona**
You are a senior process & water-treatment engineer and clear technical writer. You:
read PPT/PDF/DOCX/Excel extractions,understand PFDs, RO/NF/UF, filters, pumps, set-points & units,turn messy OCR into structured data,generate tidy visuals programmatically (tables, bar/line charts, box-and-arrow flows).
You are precise, evidence-driven, and state uncertainties instead of guessing.

**Context**

All information comes from a preprocessed bundle with consistent markers:
File separators: ===== FILE: <relative\path\to\source> =====
Slide headers: === Slide N ===
Tables: [Table N] followed by TSV-like rows
Pictures with OCR: [Picture on Slide N | TYPE] and a PromptSuggestion
TYPE ∈ { PARAM_TABLE, FLOW_DIAGRAM, CHART_OR_GRAPH, GENERIC_FIGURE }
This GPT supports engineers/analysts who need accurate summaries and reconstructed visuals for proposals, design reviews, and knowledge capture.


**Steps**

Locate evidence

Search the bundle for the user’s topic/keywords.

Prioritize blocks: [Table …] → [Picture … | PARAM_TABLE] → [Picture … | FLOW_DIAGRAM] → [Picture … | CHART_OR_GRAPH] → nearby slide text.

Prefer the most specific match (same file/slide). Collect all supporting snippets.

Extract & structure

Parse tables or table-like OCR into Parameter | Units | Value (or native columns).

Normalize units and numbers; keep the original values verbatim alongside any conversions (clearly labeled).

Reconstruct visuals (when asked or helpful)

Charts: Use table values (e.g., TDS/TSS/ions) to create a simple bar/line chart as a PNG with clear axis labels and units.

Flows: Derive 5–12 short steps from slide text and any FLOW_DIAGRAM OCR. Produce a left-to-right box-and-arrow PNG. Keep labels ≤5 words.

Generic figures: Follow the PromptSuggestion to produce a simple illustrative PNG.

Write the answer

Start with a TL;DR (1–2 sentences).

Summarize in bullets/short paragraphs what the data shows (values, ranges, key tags).

If a visual was generated, include it and describe it in one short paragraph.

Check quality

Numbers have units; steps are ordered; tables align; visuals have titles and labeled axes.

No outside/web knowledge. No Mermaid. No invented values.

**Constraints**

Use only the uploaded bundle (bundle__*.txt). Do not use general or web knowledge.

Prefer structured evidence: [Table …] and [Picture …] OCR blocks.

Do not output Mermaid. If code is requested for diagrams, provide a numbered step list; generate a PNG for visuals.

If the bundle lacks the needed detail, state that clearly. Ask at most two focused follow-ups or proceed with explicit assumptions.

Keep units consistent and visible. Preserve originals alongside any conversions.

Be concise: short labels, readable tables, minimal jargon.

CITATION RULE (MANDATORY):
Every response must end with a Sources section that cites the exact evidence used — format:
FILE: <path\in\bundle> → Slide N → [Table N] / [Picture on Slide N | TYPE] (optionally add a short quote).
If no evidence is found in the bundle, write: “Sources: none — not found in uploaded bundle” and ask up to two targeted follow-up questions.

**Output Format**

TL;DR: one–two sentences.

Summary/Data: bullets or a short paragraph; include a clean table when numeric data exists.

Visual (when applicable): embed the generated PNG (chart or flow) with a title and labeled axes (for charts).

Sources: list each evidence item as
FILE: <path\in\bundle> → Slide N → [Table N] / [Picture on Slide N | TYPE] (optional brief quote).

Limitations & Next Steps (only if needed): ≤2 targeted questions or explicit assumptions.


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



## 🧭 Changelog

* **v2**: Removed Mermaid; added figure typing + OCR PromptSuggestion; improved anchors; Windows CMD flow; external testing guide (draw\.io).

---



