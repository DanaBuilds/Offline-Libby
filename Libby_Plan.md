# Libby — Technical Implementation Plan
> Translates specifications into architectural choices, data flow, library decisions and implementation detail.
> This is the HOW — the constitution.md is the WHY, the specifications.md is the WHAT.
> Must remain aligned with constitution.md at all times.
> Living document — updated each version as the architecture evolves.

**Current version:** V13
**Last updated:** April 2026
**Spec reference:** specifications.md (SPEC-001 through SPEC-008)

---

## Constitution Compliance Check

Before any implementation decision is made, it must pass these gates from constitution.md:

| Principle | Status | Notes |
|-----------|--------|-------|
| Fully offline — no outbound calls at runtime | ✅ | All inference, search and storage is local |
| No user data leaves the machine | ✅ | No API calls, no telemetry, no logging to cloud |
| Single .py file until feature-complete | ✅ | libby_V13.py — no module split yet |
| Charcoal and rose gold design language | ✅ | Enforced in THEMES dict and report engine |
| Segoe UI font only | ✅ | All font constants use Segoe UI |
| Graceful fallbacks for optional libraries | ✅ | Pillow and watchdog wrapped in try/except |
| No hardcoded paths | ✅ | SCRIPT_DIR used throughout |
| Threading — queries never block the UI | ✅ | daemon threads for all RAG and load operations |

---

## Technical Context

| Property | Decision |
|----------|---------|
| Language | Python 3.14 |
| Platform | Windows 10/11 (64-bit), AMD Ryzen 7, 32GB RAM |
| UI Framework | tkinter + CustomTkinter |
| LLM Runtime | Ollama (local HTTP API at localhost:11434) |
| Default LLM | gpt-oss:20b |
| Vector Database | ChromaDB (local persistent client) |
| Embedding Model | all-MiniLM-L6-v2 via sentence-transformers |
| File Reading | openpyxl (read), pymupdf/fitz (PDF), built-in (txt) |
| Report Writing | xlsxwriter |
| Image Handling | Pillow (PIL) |
| File Monitoring | watchdog |
| Data Processing | pandas |
| Config Storage | JSON (libby_config.json) |
| Audit Storage | JSON (feedback_log.json) |
| Eval Storage | JSON (libby_eval_set.json) |
| Code Quality | ruff (ruff.toml) |
| Project Type | Desktop application — single file |
| Deployment Target | Windows desktop, future: Raspberry Pi / Mini PC |

---

## Architecture Overview

Libby is a single-file desktop application with five distinct layers that communicate in one direction — UI calls engine, engine never calls UI directly.

```
┌─────────────────────────────────────────────┐
│                   UI LAYER                  │
│  LibbyApp (tkinter) — tabs, bubbles, input  │
└────────────────────┬────────────────────────┘
                     │ calls
┌────────────────────▼────────────────────────┐
│               ORCHESTRATION                 │
│  Threading, tab routing, history management │
└──────┬─────────────┬───────────────┬────────┘
       │             │               │
┌──────▼──────┐ ┌────▼─────┐ ┌──────▼──────┐
│  RAG ENGINE │ │BI ENGINE │ │ EVAL ENGINE │
│  retrieve() │ │ pandas   │ │run_eval()   │
│  ask_ollama │ │ calcs    │ │             │
└──────┬──────┘ └────┬─────┘ └─────────────┘
       │             │
┌──────▼─────────────▼────────────────────────┐
│              DATA LAYER                     │
│  ChromaDB │ openpyxl │ fitz │ dataframes{} │
└─────────────────────────────────────────────┘
       │
┌──────▼──────────────────────────────────────┐
│           INFRASTRUCTURE LAYER              │
│  MCP (watchdog) │ Config │ Audit log        │
└─────────────────────────────────────────────┘
```

---

## Data Flow

### Knowledge Assistant Query
```
User types question
        ↓
_send_question() — main thread
        ↓
daemon thread starts
        ↓
V9 Context Engine — is_followup? → enrich search_q
        ↓
retrieve(search_q) → ChromaDB cosine search → top 5 chunks (dist ≤ 1.4)
        ↓
ask_ollama(question, chunks, history, model)
        ↓
Ollama HTTP POST → localhost:11434/api/generate
        ↓
response → root.after() → UI thread
        ↓
_show_answer() → _draw_libby_bubble() → feedback log entry
```

### Enterprise BI Query
```
User types data question
        ↓
get_calculations(question) → pandas → result_lines[]
        ↓
retrieve(question) → ChromaDB → chunks[]
        ↓
ask_ollama(question, chunks, calc_results, history, "ebi", model)
        ↓
response → UI thread → _draw_libby_bubble()
```

### Report Generation
```
User types "generate report..."
        ↓
_prompt_save_report() → filedialog.asksaveasfilename()
        ↓
daemon thread → generate_excel_report()
        ↓
find_best_dataframe() → keyword score → best pandas df
        ↓
xlsxwriter.Workbook() → formats → data rows → chart → save
        ↓
_show_report_result() → success message in chat
```

### MCP File Change
```
watchdog Observer monitors knowledge folder
        ↓
LibbyMCPHandler._trigger() → 2s debounce timer
        ↓
on_change_callback(event_type, filepath)
        ↓
root.after() → _mcp_notify() → chat message
        ↓
_start_backend(fresh=False) → re-index in daemon thread
```

---

## Key Technical Decisions

### Why ChromaDB
ChromaDB is the only vector database that runs fully embedded in Python with zero server process required. It persists to disk automatically, supports cosine similarity natively, and has no internet dependency after install. Alternatives like Pinecone, Weaviate or Qdrant all require network connections or separate server processes — incompatible with Libby's air-gap requirement.

### Why Ollama
Ollama is the simplest local LLM runtime available for Windows. It handles model management, quantisation and serving through a single HTTP API identical to OpenAI's format. It runs as a background service and supports all major open LLMs. llama.cpp is a lighter alternative for Raspberry Pi deployment and is noted in the roadmap.

### Why sentence-transformers (all-MiniLM-L6-v2)
all-MiniLM-L6-v2 is a 22MB embedding model that produces high quality semantic embeddings at fast inference speed on CPU. It runs entirely locally, requires no GPU, and is well suited to Libby's knowledge base sizes. It is the standard choice for offline RAG pipelines at this scale.

### Why openpyxl for reading + xlsxwriter for writing
openpyxl is the most reliable Python library for reading .xlsx files including data_only mode for formula-resolved values. However xlsxwriter has a significantly richer API for writing — supporting charts, conditional formats and advanced cell styling that openpyxl cannot match. The two libraries are not interchangeable — openpyxl reads, xlsxwriter writes.

### Why single file
Splitting into modules introduces import complexity, circular import risk and debugging difficulty during active development. Libby is a first application being built iteratively — the single file approach with clear section dividers provides all the organisational benefit of modules without the complexity cost. Module split is planned post-feature-complete.

### Why JSON for config/audit/eval
JSON is human-readable, requires no dependencies, is natively supported in Python, and is trivially editable by the user if needed. SQLite would be more robust for audit logging at scale but adds complexity — SQLite migration is planned for persistent conversation history (SPEC-010).

### Chunk Size: 1200 / Overlap: 150
Tested across a range of document types. 1200 characters provides sufficient context for multi-sentence answers while staying within Ollama's context window budget. 150 character overlap ensures no answer is split across chunk boundaries. Smaller chunks (500) lose context, larger chunks (2000+) reduce retrieval precision.

---

## Library Dependency Map

```
libby_V13.py
├── stdlib
│   ├── tkinter          — UI framework
│   ├── threading        — daemon threads for all blocking ops
│   ├── os               — path handling
│   ├── json             — config, audit, eval storage
│   ├── uuid             — unique IDs for audit entries
│   ├── re               — regex for dollar amount extraction
│   ├── shutil           — file operations
│   ├── gc               — garbage collection on re-index
│   └── datetime         — timestamps
│
├── required (pip install)
│   ├── chromadb         — vector database
│   ├── sentence-transformers — embeddings (all-MiniLM-L6-v2)
│   ├── requests         — HTTP calls to Ollama API
│   ├── openpyxl         — reading .xlsx files
│   ├── xlsxwriter       — writing/exporting Excel reports
│   ├── pymupdf (fitz)   — reading .pdf files
│   └── pandas           — BI calculations and dataframe ops
│
└── optional (graceful fallback if missing)
    ├── Pillow (PIL)     — splash screen + sidebar logo
    └── watchdog         — MCP live file monitoring
```

---

## File & Folder Structure

```
C:\Libby\
├── libby_V13.py          — main application (single file)
├── libby_config.json     — auto-generated user settings
├── feedback_log.json     — auto-generated audit trail
├── libby_eval_set.json   — auto-generated evaluation test cases
├── ruff.toml             — linting configuration
├── constitution.md       — project principles (Spec Kit)
├── specifications.md     — feature specifications (Spec Kit)
├── plan.md               — this file (Spec Kit)
├── README.md             — setup guide and documentation
└── .gitignore            — excludes db, config and data files

C:\[Knowledge Folder]\
└── libby_db\             — ChromaDB persistent storage (auto-created)
```

---

## Data Models

### libby_config.json
```json
{
  "base_path": "C:\\RAG Test",
  "theme": "dark",
  "mode": "knowledge",
  "company_name": "Libby Intelligence",
  "ollama_model": "gpt-oss:20b"
}
```

### feedback_log.json (one entry per interaction)
```json
{
  "id": "uuid-string",
  "timestamp": "2026-04-03T21:00:00",
  "tab": "knowledge",
  "question": "What is...",
  "answer": "Based on the documents...",
  "sources": ["filename.pdf"],
  "rating": "positive"
}
```

### libby_eval_set.json (array of test cases)
```json
[
  {
    "question": "What is the reorder level for SKU-001?",
    "expected": "50 units"
  }
]
```

### ChromaDB chunk metadata
```json
{
  "source": "C:\\RAG Test\\file.pdf",
  "filename": "file.pdf",
  "file_type": "pdf",
  "chunk_index": 0
}
```

---

## Performance Constraints

| Constraint | Target | Current Status |
|-----------|--------|---------------|
| First launch index time | < 60 seconds for 50 files | ✅ Met |
| Query response time | < 30 seconds end to end | ✅ Met (model dependent) |
| UI thread blocking | Never | ✅ All ops in daemon threads |
| RAM usage at rest | < 4GB | ✅ Met on 32GB system |
| Chunk cosine distance threshold | ≤ 1.4 | ✅ Enforced in retrieve() |
| MCP debounce delay | 2 seconds | ✅ Enforced in LibbyMCPHandler |
| Conversation history stored | 20 exchanges | ✅ Enforced in _send_question() |
| Context passed to LLM | 4 exchanges | ✅ Enforced in ask_ollama() |

---

## Error Handling Strategy

| Scenario | Handling |
|---------|---------|
| Ollama not running | Returns "⚠ Ollama not running. Open CMD and type: ollama serve" |
| No chunks found | Returns "I couldn't find anything relevant. Try rephrasing." |
| EBI no data | Returns "Not found." |
| File unreadable | Prints error to console, skips file, continues indexing |
| Pillow not installed | PILLOW_AVAILABLE = False, splash and logo skipped silently |
| watchdog not installed | WATCHDOG_AVAILABLE = False, MCP disabled silently |
| Report save cancelled | Returns "Report cancelled." message in chat |
| ChromaDB corrupted | Delete libby_db folder and restart — documented in README |
| Settings invalid path | messagebox.showerror() — settings not saved |

---

## Planned Technical Changes (Roadmap)

| Spec | Feature | Technical Approach |
|------|---------|-------------------|
| SPEC-009 | OCR for scanned PDFs | pytesseract + Tesseract-OCR — detect image-only pages via fitz, hand off to pytesseract |
| SPEC-010 | Persistent conversation history | sqlite3 (stdlib) — conversations table, per-tab storage, load on startup |
| SPEC-011 | PyInstaller packaging | .spec file with hidden imports for chromadb, sentence-transformers, fitz |
| SPEC-012 | First-run setup wizard | tkinter Toplevel — detect missing libby_config.json on launch |
| SPEC-013 | Encryption layer | cryptography (Fernet) — encrypt feedback_log.json and config at rest |
| SPEC-014 | Flashcard / quiz mode | New 4th tab — pull chunks from ChromaDB, present as Q&A cards |
| SPEC-015 | Raspberry Pi deployment | llama.cpp instead of Ollama — ARM compatible, lower RAM footprint |
| SPEC-016 | .docx support | python-docx — add to file reader, extend SUPPORTED_EXTENSIONS |
| SPEC-017 | Cloud / ERP connector | Connection-optional architecture — local first, sync when connected |

---

*Technical plan established at V13.*
*All decisions must comply with constitution.md.*
*Created by DanaBuilds — github.com/DanaBuilds/Offline-Libby*