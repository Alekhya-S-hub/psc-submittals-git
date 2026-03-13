# Submittals Agent

**A hybrid automation system for extracting submittal requirements from construction specification documents.**

The Submittals Agent combines deterministic document parsing with selective LLM reasoning and human oversight to extract, structure, and organize submittal requirements from CSI MasterFormat construction specifications. It replaces a manual process that typically requires 4 engineers × 4 hours per document (~16 hours) with an automated pipeline that completes in under 5 minutes, requiring less than 1 hour of human review.

---

## Problem

Construction specifications are contractual documents that define material standards, construction procedures, and submittal obligations across hundreds of sections organized by [CSI MasterFormat](https://www.csiresources.org/standards/masterformat). A typical spec book is 800–1,200 pages. Project engineers must manually:

1. Navigate the Table of Contents to identify relevant sections
2. Locate submittal subsections (e.g., `1.03 SUBMITTALS`, `1.04 ACTION SUBMITTALS`) within each section
3. Extract obligation text with its full hierarchical context (Part → Subsection → Paragraph → Subparagraph)
4. Compile a submittal log listing every required submission
5. Generate transmittal cover sheets and folder structures for the review cycle

This process is error-prone, repetitive, and time-consuming. Generic LLM-based extraction lacks the domain-specific grounding to interpret conditional obligations, deontic phrasing (*shall*, *may*, *if required*), and cross-sectional dependencies reliably.

---

## Architecture Overview

The system decomposes the task into three layers:

```
┌─────────────────────────────────────────────────────────┐
│                    FastAPI Service                       │
│                     (app.py)                             │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  1. DETERMINISTIC PARSING (extractor.py)                │
│     ├── PDF text extraction (PyMuPDF)                   │
│     ├── TOC detection & section boundary resolution     │
│     ├── Submittal keyword matching                      │
│     ├── 10-level CSI hierarchy parsing                  │
│     └── Structured Excel/folder generation              │
│                                                         │
│  2. SELECTIVE LLM REASONING (GPT-4o-mini)               │
│     └── Project metadata extraction from first 10 pages │
│        (project name, engineer, contractor, owner)      │
│                                                         │
│  3. HUMAN REVIEW                                        │
│     └── All outputs require engineer review/approval    │
│        before use in project workflows                  │
│                                                         │
└─────────────────────────────────────────────────────────┘
```

The LLM is used **only** for extracting unstructured project metadata (names, addresses, project numbers) from cover pages — a task where deterministic parsing is unreliable due to inconsistent formatting. All submittal extraction logic is deterministic.

---

## Technical Details

### Document Parsing Pipeline

**Full-text extraction** is performed once at initialization using PyMuPDF's native text extraction (`page.get_text("text", sort=True)`), caching the entire document in memory. All subsequent operations search this cached string rather than re-reading the PDF.

**TOC detection** scans the first 150 pages for `TABLE OF CONTENTS` headers, then selects the TOC instance with the highest density of CSI section numbers (pattern: `\d{2}\s+\d{2}\s+\d{2}` or `\d{5,6}`). This handles spec books with multiple TOCs (e.g., a bid-document TOC followed by the technical specifications TOC).

**Section boundary resolution** locates each section in the cached text by matching `SECTION <number>` headers at line boundaries (case-sensitive to avoid false matches with inline references), then extracts all text up to the next `SECTION` header.

### Regex Patterns

The extractor uses pre-compiled regex patterns for performance. Key patterns include:

**Section header detection:**
```python
# Matches "SECTION 011100" or "SECTION 01 11 00"
r'SECTION\s+(\d{2}\s+\d{2}\s+\d{2}|\d{5,6}(?:\.\d+)?)'
```

**TOC line parsing:**
```python
# Matches "011100 Summary" or "01 11 00 Summary" with optional trailing page numbers
r'^(\d{2}\s+\d{2}\s+\d{2}|\d{5,6}(?:\.\d+)?)\s+(.+?)(?:\s+\d+\s*)?$'
```

**Submittal subsection identification** uses keyword matching against a defined list:
```python
SUBMITTAL_KEYWORDS = [
    "ACTION SUBMITTALS", "INFORMATIONAL SUBMITTALS", "CLOSEOUT SUBMITTALS",
    "SHOP DRAWING SUBMITTALS", "AS-BUILT SUBMITTALS", "QUALITY ASSURANCE SUBMITTALS",
    "SUBMITTAL REQUIREMENTS", "SUBMITTAL SCHEDULE", "FORM OF SUBMITTALS",
    "RECORD DOCUMENT SUBMITTALS", "SUBMITTALS"
]
```

### CSI Hierarchy Parsing (10 Levels)

Construction specifications follow a standardized hierarchical numbering scheme. The extractor recognizes all 10 levels:

| Level | Pattern | Example |
|-------|---------|---------|
| 1 | `PART N` | `PART 1 - GENERAL` |
| 2 | `N.NN` | `1.03 SUBMITTALS` |
| 3 | `A.` | `A. Product Data` |
| 4 | `1.` | `1. Submit manufacturer certificates` |
| 5 | `a.` | `a. Include test reports` |
| 6 | `1)` | `1) Compressive strength data` |
| 7 | `a)` | `a) 28-day test results` |
| 8 | `(a)` | `(a) Per ASTM C39` |
| 9 | `(1)` | `(1) Minimum 3 specimens` |
| 10 | `i.` | `i. Cured under lab conditions` |

Each pattern is pre-compiled and applied in order. Continuation lines (text that wraps to the next line without a new hierarchy marker) are merged with their parent item.

### Multi-line Heading Resolution

PDF text extraction frequently splits headings across lines. For example:
```
1.04
SUBMITTALS
```
The extractor detects orphaned subsection numbers (`^\d+\.\d+$`) and hierarchy markers (`^[A-Z]\.$`) and merges them with the following line before hierarchy parsing.

### LLM Usage (Project Metadata Only)

The LLM (GPT-4o-mini, temperature=0) is invoked **once per document** to extract project metadata from the first 10 pages. The prompt is constrained to return a fixed JSON schema:

```json
{
  "project_name": "",
  "ccua_project_number": "",
  "pscc_job_number": "",
  "engineer_name": "",
  "engineer_address": "",
  "contractor_name": "",
  "contractor_address": "",
  "owner_name": "",
  "owner_address": "",
  "prepared_by": ""
}
```

Input text is truncated to 12,000 characters. The system prompt restricts the model to JSON-only output. This metadata is used to populate transmittal templates — it does not influence submittal extraction.

---

## API Endpoints

| Endpoint | Method | Description | Output |
|----------|--------|-------------|--------|
| `/extract-submittals-sections` | POST | Extract submittal content with full hierarchy | Excel workbook (one sheet per section) |
| `/extract-submittals-log` | POST | Generate submittal log summary | Excel workbook (log format) |
| `/extract-project-info` | POST | Extract project metadata via LLM | JSON |
| `/create-submittal-structure` | POST | Generate complete folder structure with transmittals and cover sheets | ZIP archive |
| `/health` | GET | Health check | JSON |

All PDF endpoints accept base64-encoded PDF content in the request body.

---

## Outputs

For each specification document, the system generates:

- **Submittal Sections Workbook** — One Excel sheet per specification section containing submittal requirements, with columns for each hierarchy level (Levels 1–10)
- **Submittal Log** — Summary Excel listing all sections with submittal requirements, formatted for project tracking
- **Folder Structure** — Per-section directories with subdirectories for the review cycle (`From Vendor → To Engineer → From Engineer → Final Approved`), including revision tracking folders
- **Transmittals & Cover Sheets** — Pre-filled Word and Excel templates with project metadata, section numbers, and submittal identifiers

---

## Installation

### Prerequisites

- Python 3.10+
- An OpenAI API key (for project metadata extraction only)

### Setup

```bash
git clone https://github.com/<your-username>/submittals-agent.git
cd submittals-agent

pip install -r requirements.txt

# Configure environment
cp .env.example .env
# Add your OPENAI_API_KEY to .env
```

### Dependencies

```
fastapi
uvicorn
pymupdf
openpyxl
python-docx
openai
python-dotenv
pydantic
```

### Run

```bash
uvicorn app:app --host 0.0.0.0 --port 8000
```

API documentation available at `http://localhost:8000/docs`.

### Standalone Extraction (No API)

```bash
python extractor.py path/to/specification.pdf
# Outputs: test_sections.xlsx, test_log.xlsx
```

---

## Evaluation

Evaluated on 20 construction specification documents (800–1,200 pages each, avg. 1,000 pages) from real projects:

| Metric | Value |
|--------|-------|
| Precision | 96.2% |
| Recall | 92.5% |
| F1-Score | 94.3% |
| Processing time per document | ~3 minutes |
| Human review time per document | <1 hour |
| Time reduction vs. manual | 94% |
| Cost reduction vs. manual | 93% |

False negatives primarily occur in sections that deviate from standard CSI formatting. False positives arise from informational references to submittals rather than contractual obligations. Both are caught during human review.

---

## Deployment

The system is deployed as an Azure Web App (premium tier) and orchestrated via Microsoft Power Automate for integration with SharePoint/OneDrive workflows. A Microsoft Copilot Studio agent provides a conversational interface for project teams.

---

## Templates

Place the following templates in the `templates/` directory:

- `SubmittalLog.xlsx` — Template for the submittal log output
- `Coverpage_New.docx` — Template for submittal cover sheets (with `{{PLACEHOLDER}}` fields)
- `Transmittal.xlsx` — Template for transmittal forms (with `{{PLACEHOLDER}}` fields)

Supported placeholders: `{{PROJECT_NAME}}`, `{{CCUA_PROJECT_NO}}`, `{{PSCC_JOB_NO}}`, `{{ENGINEER_NAME}}`, `{{ENGINEER_ADDRESS}}`, `{{CONTRACTOR_NAME}}`, `{{CONTRACTOR_ADDRESS}}`, `{{OWNER_NAME}}`, `{{OWNER_ADDRESS}}`, `{{PREPARED_BY}}`, `{{SECTION_NUMBER}}`, `{{SUBMITTAL_TITLE}}`, `{{SUBMITTAL_NUMBER}}`, `{{DATE}}`

---

## Limitations

- Assumes specifications follow CSI MasterFormat conventions. Non-standard section numbering may reduce recall.
- Scanned (image-only) PDFs require OCR (`use_ocr=True`), which is slower and less accurate.
- The LLM component (project metadata) requires an OpenAI API key and network access.
- Evaluated on specifications from a single geographic region (North America) and a limited set of contractors.
- Footer filtering includes hardcoded patterns for specific firms — these should be generalized for broader use.

---

## License

[MIT License](LICENSE)
