# Libby — Offline AI Knowledge Assistant

> Libby is a fully offline, air-gapped RAG assistant that turns local document folders into a private, conversational knowledge base. It runs entirely on your local machine with zero cloud dependencies.


![Python](https://img.shields.io/badge/Python-3.14-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![Version](https://img.shields.io/badge/Version-V16-brightgreen)
![Offline](https://img.shields.io/badge/Mode-Fully%20Offline-orange)
![Status](https://img.shields.io/badge/Status-Active-success)

---

## What is Libby?

Libby is a fully offline, air-gapped AI assistant built in Python. She answers questions from a local knowledge base using RAG (Retrieval Augmented Generation). No internet required after initial setup.

Libby uses a two-component local architecture to parse, index, and query documents completely offline:

As of V14:

| Component | File | Role |
|---|---|---|
| **Libby** | `libby_V16.py` | The AI chat interface — asks questions, retrieves answers |
| **Up** | `Up_V2.py` | The document mirror engine — converts source files into clean `.txt` copies |

Up runs automatically in the background whenever Libby starts. You never need to run it manually.

---

## What's New in V16

- **Multiple knowledge source folders** — add as many source folders as you like via Settings
- **Shared mirror** — all folders sync into a single `_libby_mirror/` next to `libby_V16.py`, Libby searches across everything in one pass
- **Scrollable Settings panel** — Settings window now scrolls cleanly regardless of screen size
- **MCP feedback loop fix** — watcher pauses during Up sync to prevent repeated re-indexing

---

## Recent Version Highlights

**V15**
- Single **Ask Libby** tab — Knowledge Assistant and Enterprise BI merged; Libby auto-detects the question mode internally
- Scrollable 15-line message bubbles for long answers
- Sidebar logo and splash screen fully theme-aware (dark and light)

**V14**
- **Up integration** — Libby reads only from `_libby_mirror/`, Up handles all file conversion
- `source_folder` config key — Libby derives the mirror path automatically
- MCP watcher moved to monitor the mirror folder

---

## Architecture

```
Source Folders (your documents)
  C:\RAG Test\          C:\Knowledge\
       │                      │
       └──────────┬───────────┘
                  ▼
            Up_V2.py (sync engine)
            Converts .pdf .xlsx .docx .txt → clean .txt
                  │
                  ▼
         C:\Libby\_libby_mirror\
           (shared mirror — all folders merged)
                  │
                  ▼
            libby_V16.py
            ChromaDB indexes .txt files
            Ollama answers questions
```

Originals are **never touched**. Up maintains a SQLite index (`up_index.db`) tracking every file's hash, sync status, and full section → chunk → sentence hierarchy.

---

## System Requirements

- Windows 10 or 11 (64-bit)
- AMD Ryzen 7 or equivalent (recommended)
- 16GB RAM minimum (32GB recommended for larger knowledge bases)
- 20GB free disk space
- Internet connection for initial setup only

---

## Step 1 — Install Python

1. Go to `https://www.python.org/downloads/`
2. Download the latest **Python 3.x** version for Windows (64-bit)
3. Run the installer
4. **Important:** Check **"Add Python to PATH"** before clicking Install
5. Click **Install Now**

Verify it worked — open PowerShell and type:
```
python --version
```
You should see something like `Python 3.14.0`

---

## Step 2 — Install Required Libraries

Open PowerShell and run:
```
pip install chromadb sentence-transformers openpyxl pymupdf requests pandas xlsxwriter Pillow watchdog python-docx
```

This may take a few minutes. Wait for it to finish before moving on.

> **Note:** `python-docx` is new in V14 — required for Up to convert `.docx` files.

---

## Step 3 — Install Ollama

Ollama runs the AI model locally on your machine.

1. Go to `https://ollama.com/download`
2. Click **Download for Windows**
3. Run the installer like any normal program

Verify it worked:
```
ollama --version
```

---

## Step 4 — Download an AI Model

```
ollama pull gpt-oss:20b
```

**Supported models:**

| Model | Size | Notes |
|---|---|---|
| `gpt-oss:20b` | ~12GB | Recommended — most capable, requires 32GB RAM |
| `llama3.2` | ~2GB | Good balance of speed and quality |
| `llama3.1:8b` | ~5GB | Good mid-range option |
| `mistral` | ~4GB | Fast and reliable |
| `phi3:mini` | ~2GB | Lightweight — good for lower RAM machines |

You can change the active model at any time in Libby's Settings panel.

---

## Step 5 — Add Your Knowledge Files

Place your files in any folder on your machine. Libby supports multiple source folders — you can add them all in Settings.

**Supported file types (converted by Up):**

| Type | Best For |
|---|---|
| `.txt` | Notes, guides, manuals, written knowledge |
| `.xlsx` | Business data, spreadsheets, sales, inventory |
| `.pdf` | Reports, documents, reference material |
| `.docx` | Word documents, policies, procedures |

**Tips for best results:**
- Use clear descriptive filenames — Libby cites filenames in her answers
- Organise files into subfolders by topic — subfolder names become searchable tags
- Excel files work best with clear column headers in row 1
- PDFs must be text-based (not scanned images)

---

## Step 6 — Run Libby

In PowerShell, navigate to your Libby folder and run:
```
cd C:\Libby
python libby_V16.py
```

On first launch:
1. Up syncs all configured source folders into `_libby_mirror/`
2. Libby indexes the mirror into ChromaDB
3. The splash screen appears while loading

Subsequent launches are much faster — Up only processes changed files.

---

## Step 7 — Configure Your Settings

Click **Settings** in the top right corner.

| Setting | What it does |
|---|---|
| **Knowledge Source Folders** | Add/remove source folders — Up syncs all of them |
| **Theme** | Switch between Dark Mode (charcoal) and Light Mode (warm linen) |
| **Company Name** | Personalise the header and exported reports |
| **AI Model** | Switch between installed Ollama LLMs |

Click **Save & Apply** — Up syncs immediately, Libby reloads from the mirror.

To add a new source folder: click **+ Add Folder**, select it, click Save & Apply. Up converts it automatically.

---

## How to Use Libby

### Ask Libby Tab

Ask anything in plain English — Libby automatically detects whether your question is a knowledge question or a data question and switches mode internally.

**Knowledge questions:**
```
How do I treat a second degree burn?
What foods have the longest shelf life?
How long should I boil water to purify it?
```

**Data questions (auto-detected):**
```
What is the total salary budget?
Which product has the highest profit margin?
What is the average revenue per sales rep?
```

**Generating reports:**

Start your message with **"generate report"**:
```
generate report showing sales by region highest to lowest
generate report of inventory items lowest to highest stock
generate report showing employees by salary highest to lowest
```

A save dialog opens. Libby builds a fully formatted Excel report with:
- Company header in charcoal and rose gold branding
- Sorted and grouped data rows with alternating colours
- Currency formatting on financial columns
- Totals row
- Auto-generated rose gold bar chart
- Confidentiality footer

### Feedback

Every Libby response has thumbs up / thumbs down buttons. Ratings are saved to `feedback_log.json` as an audit trail.

---

## Up — Document Mirror Engine

Up (`Up_V2.py`) runs automatically as part of Libby. It can also be run standalone from the terminal:

```
python Up_V2.py "C:\path\to\your\folder"
```

Output example:
```
Up  —  syncing: C:\RAG Test
Mirror: C:\RAG Test\_libby_mirror

  ➕  [added]  budget.xlsx
  🔄  [updated]  notes.txt

✅  Added:    1
🔄  Updated:  1
📦  Archived: 0
⏭   Skipped:  11

Index: 12 docs  •  54 sections  •  183 chunks  •  1,204 sentences
```

**How Up works:**
- Scans source folders for `.txt`, `.pdf`, `.xlsx`, `.docx` files
- Converts each to a clean `.txt` with a provenance header
- Uses SHA-256 hashing — only re-converts files that actually changed
- Deleted source files → mirror copy moved to `_archive/` with a timestamp
- Builds a SQLite index with a full document → section → chunk → sentence hierarchy
- Infers tags from filenames and subfolder names automatically

---

## MCP — Live File Monitoring

Libby includes a Model Context Protocol (MCP) layer powered by `watchdog` that monitors `_libby_mirror/` in real time:

- Detects when Up writes new or updated `.txt` files to the mirror
- Auto re-indexes ChromaDB without restarting Libby
- 8 second debounce prevents duplicate events
- Watcher pauses during Up sync to prevent feedback loops
- Graceful fallback if watchdog is not installed

---

## Evaluation Tab

The Evaluation tab is a developer tool for testing Libby's accuracy:

- Add question and expected answer pairs
- Run the full test suite silently through the RAG pipeline
- Keyword matching scorer checks expected answer keywords appear in Libby's response
- Per-question ✅ ❌ display with live progress counter
- Final score with colour coding — green 70%+, amber 40%+, red below 40%
- Test cases saved persistently in `libby_eval_set.json`

> **Note:** Make sure your eval questions match the content in your loaded knowledge base.

---

## Privacy & Air-Gap

Libby is designed to work **completely offline:**

- No data sent to the internet
- No cloud APIs used
- All knowledge stays on your machine
- Ollama and ChromaDB run entirely locally
- Generated reports never leave your computer unless you share them
- Compatible with Faraday bag deployment for true signal isolation

After initial setup Libby never needs an internet connection.

---

## Troubleshooting

**Libby shows "0 chunks" or files are missing:**
- Open Settings, confirm your source folders are listed, click Save & Apply
- Delete `C:\Libby\_libby_mirror\libby_db\` and restart Libby to force a full rebuild
- Check that your files are a supported type (`.txt`, `.xlsx`, `.pdf`, `.docx`)

**"Ollama not running" error:**
- Open a new PowerShell window and type: `ollama serve`
- Leave that window open and restart Libby

**Excel file not loading:**
- Open the file in Excel and re-save as `.xlsx`
- Make sure the file is not open in another program when Libby starts
- Check that row 1 contains column headers

**MCP keeps re-indexing in a loop:**
- This was fixed in V16 — update to the latest version

**ModuleNotFoundError on launch:**
- Run: `pip install xlsxwriter Pillow watchdog python-docx`
- All libraries must be installed before running Libby

**ruff not recognised as a command:**
- Use `python -m ruff check libby_V16.py` instead
- Make sure you are in the `C:\Libby` directory first

---

## Requirements Summary

| Requirement | Version |
|---|---|
| Python | 3.10 or higher |
| chromadb | latest |
| sentence-transformers | latest |
| openpyxl | latest |
| pymupdf (fitz) | latest |
| requests | latest |
| pandas | latest |
| xlsxwriter | latest |
| Pillow | latest |
| watchdog | latest |
| python-docx | latest |
| Ollama | latest |
| Recommended LLM | gpt-oss:20b (~12GB) |

---

## Auto-Generated Files

Libby and Up create these automatically — you never need to edit them:

| File / Folder | Purpose |
|---|---|
| `libby_config.json` | Saves your settings between sessions |
| `feedback_log.json` | Audit trail of every interaction and rating |
| `libby_eval_set.json` | Saved evaluation test cases |
| `_libby_mirror/` | Shared mirror — all source folders merged here |
| `_libby_mirror/libby_db/` | ChromaDB vector database |
| `source_folder/_libby_mirror/` | Per-folder mirror managed by Up |
| `source_folder/_libby_mirror/up_index.db` | Up's SQLite document index |
| `source_folder/_libby_mirror/_archive/` | Archived mirrors of deleted source files |

---

## Roadmap

- [ ] Tune evaluation scorer to reduce false negatives
- [ ] Package as `.exe` with PyInstaller — one-click launch, no Python required
- [ ] First-run setup wizard
- [ ] Deploy to offline hardware (Raspberry Pi / Mini PC)
- [ ] Solar power integration for true off-grid use
- [ ] OCR support for scanned PDFs (pytesseract)
- [ ] Persistent conversation history between sessions
- [ ] Flashcard and quiz mode for personal knowledge base
- [ ] Cybersecurity layer — encryption (Fernet), bandit scanning
- [ ] Cloud / ERP connector (connection-optional architecture)
- [ ] Outlook `.msg` file support (extract-msg)
- [ ] Wing-based topic routing (V17)

---

## Version History

| Version | Highlights |
|---|---|
| V1 | Basic RAG, CLI, Mistral |
| V2 | Tkinter UI, phi3:mini, strict mode |
| V3 | Dark/light themes, PDF support |
| V4 | Enterprise BI, pandas calculations |
| V5 | Ultra-concise mode, professional UI |
| V6 | Smart report generation, conversation memory |
| V7 | Universal app, dual tab system, settings panel, folder switching |
| V8 | Audit log, feedback_log.json, thumbs up/down rating system |
| V9 | Conversational Context Engine, follow-up detection, enriched queries |
| V10 | MCP file watcher, live auto re-indexing, watchdog integration |
| V11 | Evaluation tab, keyword scoring, libby_eval_set.json |
| V12 | gpt-oss:20b LLM upgrade, config refinements |
| V13 | Pillow splash screen + sidebar logo, xlsxwriter reports with charts, ruff linting |
| V14 | Up integration, mirror architecture, source_folder config, MCP watches mirror |
| V15 | Single Ask Libby tab, auto mode detection, scrollable bubbles, theme-aware logo |
| **V16** | **Multi-folder support, shared mirror, scrollable Settings, MCP feedback loop fix** |

---

*Built with Python, ChromaDB, Ollama, sentence-transformers, Pillow, xlsxwriter and python-docx*
*Designed for offline, air-gapped environments*
*Created by DanaBuilds — github.com/DanaBuilds/Offline-Libby*
