# 📚 Libby — Offline AI Knowledge Assistant

> *"Libby is a local AI framework that turns any folder of documents into a conversational knowledge base — deployable for any domain, customizable for any user."*

![Python](https://img.shields.io/badge/Python-3.14-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![Version](https://img.shields.io/badge/Version-V13-brightgreen)
![Offline](https://img.shields.io/badge/Mode-Fully%20Offline-orange)
![Status](https://img.shields.io/badge/Status-Active-success)

---

## What is Libby?

Libby is a fully offline, air-gapped AI assistant built in Python. She answers questions from a local knowledge base of `.txt`, `.xlsx` and `.pdf` files using RAG (Retrieval Augmented Generation). No internet required after initial setup.

Libby is designed to be deployed anywhere — a desktop PC, a mini PC, a Raspberry Pi, or a solar-powered handheld device — with zero cloud dependency and zero data leaving your machine.

---

## ✨ What's New in V13

- **Pillow splash screen** — branded charcoal and rose gold startup screen on launch
- **Pillow sidebar logo** — rose gold book icon rendered at the top of the sidebar
- **xlsxwriter report engine** — Excel reports upgraded with richer styling and auto-generated rose gold bar charts
- **ruff linting** — `ruff.toml` config added for code quality scanning
- **Input focus fix** — resolved splash screen grab that was blocking the input box

---

## 🖥️ System Requirements

- Windows 10 or 11 (64-bit)
- AMD Ryzen 7 or equivalent (recommended)
- 16GB RAM minimum (32GB recommended for larger knowledge bases)
- 20GB free disk space
- Internet connection for initial setup only

---

## 🗂️ Folder Structure

Create a folder anywhere on your computer for your knowledge base. Libby will find all supported files automatically, no matter how deep the subfolders go.

```
C:\My Knowledge\
  ├── Health\
  │     ├── first_aid_basics.txt
  │     └── medication_guide.pdf
  ├── Business\
  │     ├── sales_data.xlsx
  │     └── inventory.xlsx
  └── Survival\
        └── water_purification.txt
```

Libby auto-creates a `libby_db` folder inside your knowledge folder to store her vector database. You never need to touch it.

---

## ⚙️ Step 1 — Install Python

1. Go to `https://www.python.org/downloads/`
2. Download the latest **Python 3.x** version for Windows (64-bit)
3. Run the installer
4. ⚠️ **Important:** Check **"Add Python to PATH"** before clicking Install
5. Click **Install Now**

Verify it worked — open PowerShell and type:
```
python --version
```
You should see something like `Python 3.14.0`

---

## 📦 Step 2 — Install Required Libraries

Open PowerShell and run:
```
pip install chromadb sentence-transformers openpyxl pymupdf requests pandas xlsxwriter Pillow watchdog
```

This may take a few minutes. Wait for it to finish before moving on.

---

## 🤖 Step 3 — Install Ollama

Ollama runs the AI model locally on your machine.

1. Go to `https://ollama.com/download`
2. Click **Download for Windows**
3. Run the installer like any normal program

Verify it worked:
```
ollama --version
```

---

## 🧠 Step 4 — Download an AI Model

Libby works with any Ollama-compatible LLM (Large Language Model). The current recommended model is `llama3.2`.

```
ollama pull llama3.2
```

This downloads approximately 2GB. This is the **last time you need internet** for Libby to work.

**Other supported models:**

| Model | Size | Notes |
|---|---|---|
| `llama3.2` | ~2GB | Recommended — great balance of speed and quality |
| `gpt-oss:20b` | ~12GB | Most capable — requires 32GB RAM |
| `llama3.1:8b` | ~5GB | Good mid-range option |
| `mistral` | ~4GB | Fast and reliable |
| `phi3:mini` | ~2GB | Lightweight — good for lower RAM machines |

You can change the active model at any time in Libby's Settings panel without reinstalling anything.

---

## 📄 Step 5 — Add Your Knowledge Files

Place your files inside your knowledge folder in any structure you like.

**Supported file types:**

| Type | Icon | Best For |
|---|---|---|
| `.txt` | 📄 | Notes, guides, manuals, written knowledge |
| `.xlsx` | 📊 | Business data, spreadsheets, sales, inventory |
| `.pdf` | 📕 | Reports, documents, reference material |

**Tips for best results:**
- Use clear descriptive filenames — Libby cites filenames in her answers
- Organise files into subfolders by topic
- Excel files work best with clear column headers in row 1
- PDFs must be text-based (not scanned images) — OCR support coming in a future version

---

## 🚀 Step 6 — Run Libby

In PowerShell, navigate to your Libby folder and run:
```
cd C:\Libby
python libby_V13.py
```

On first launch Libby will index all your files and display a splash screen. Indexing may take a minute depending on how many files you have. Subsequent launches are much faster.

---

## ⚙️ Step 7 — Configure Your Settings

Click **⚙ Settings** in the top right corner to configure:

| Setting | What it does |
|---|---|
| **Knowledge Folder** | Point Libby at your files — click Browse to select any folder |
| **Theme** | Switch between Dark Mode (charcoal) and Light Mode (warm linen) |
| **Company Name** | Personalise the header and exported reports |
| **AI Model** | Switch between installed Ollama LLMs |

Click **Save & Apply** — changes take effect immediately, no restart needed.

When you change the knowledge folder Libby automatically clears her database and rebuilds from the new location.

---

## 💬 How to Use Libby

### 📚 Knowledge Assistant Tab

Ask questions in plain English from your loaded documents:

```
How do I treat a second degree burn?
What foods have the longest shelf life?
How long should I boil water to purify it?
```

Libby answers using only what is in your knowledge base — never guessing or using outside information.

### 📊 Enterprise BI Tab

Ask data questions from your Excel files:

```
What is the total salary budget?
Which product has the highest profit margin?
Which customers are on hold?
What is the average revenue per sales rep?
```

Libby performs real calculations using your actual data and cites the source file.

### 📊 Generating Reports

In the Enterprise BI tab, start your message with **"generate report"**:

```
generate report showing sales by region highest to lowest
generate report of inventory items lowest to highest stock
generate report of active customers by outstanding balance
generate report showing employees by salary highest to lowest
```

A save dialog opens — choose where to save your Excel report. Libby builds a fully formatted report with:
- Company header in charcoal and rose gold branding
- Sorted and grouped data rows with alternating row colours
- Currency formatting on financial columns
- Totals row
- Auto-generated rose gold bar chart
- Confidentiality footer

### 👍👎 Feedback

Every Libby response has thumbs up / thumbs down buttons. Ratings are saved to `feedback_log.json` as an audit trail.

---

## 🗂️ Tab System

Libby has three tabs, each with independent state:

| Tab | Purpose |
|---|---|
| 📚 Knowledge Assistant | Document Q&A from your knowledge base |
| 📊 Enterprise BI | Data questions, calculations and report generation |
| 🧪 Evaluation | Developer tool for accuracy testing — not end user facing |

- Each tab keeps its own separate conversation history
- A dot indicator (●) appears on inactive tabs that have conversation history
- The **Clear** button only clears the active tab
- History resets when you close Libby (persistent history coming in a future version)

---

## 🔁 MCP — Live File Monitoring

Libby includes a Model Context Protocol (MCP) layer powered by `watchdog` that monitors your knowledge folder in real time:

- Detects when files are **added, modified, deleted or renamed**
- Auto re-indexes the knowledge base without restarting
- 2 second debounce prevents duplicate events
- Sends a notification message in the chat when files change
- Restarts automatically when you change the knowledge folder in Settings
- Graceful fallback if watchdog is not installed

To add new files to Libby on a deployment device — simply copy files via USB stick into the knowledge folder. MCP handles the rest automatically.

---

## 🧪 Evaluation Function

The Evaluation tab is a developer tool for testing Libby's accuracy:

- Add question and expected answer pairs
- Run the full test suite silently through the RAG pipeline
- Keyword matching scorer checks expected answer keywords appear in Libby's response
- Per-question ✅ ❌ display with live progress counter
- Final score with colour coding — green 70%+, amber 40%+, red below 40%
- Test cases saved persistently in `libby_eval_set.json`

---

## 🔒 Privacy & Air-Gap

Libby is designed to work **completely offline:**

- ✅ No data sent to the internet
- ✅ No cloud APIs used
- ✅ All knowledge stays on your machine
- ✅ Ollama and ChromaDB run entirely locally
- ✅ Generated reports never leave your computer unless you share them
- ✅ Compatible with Faraday bag deployment for true signal isolation

After initial setup Libby never needs an internet connection.

---

## 🛠️ Troubleshooting

**Libby shows "0 chunks" or files are missing:**
- Make sure your files are inside the selected knowledge folder
- Delete the `libby_db` folder inside your knowledge folder and restart
- Check that your files are `.txt`, `.xlsx` or `.pdf` format

**"Ollama not running" error:**
- Open a new PowerShell window and type: `ollama serve`
- Leave that window open and restart Libby

**Excel file not loading:**
- Open the file in Excel and re-save as `.xlsx`
- Make sure the file is not open in another program when Libby starts
- Check that row 1 contains column headers

**Input box not responding after startup:**
- Wait for the splash screen to fully close before clicking
- If still frozen restart Libby — this was fixed in V13

**ModuleNotFoundError on launch:**
- Run: `pip install xlsxwriter Pillow watchdog`
- All libraries must be installed before running Libby

**ruff not recognised as a command:**
- Use `python -m ruff check libby_V13.py` instead
- Make sure you are in the `C:\Libby` directory first

---

## 📋 Requirements Summary

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
| Ollama | latest |
| Recommended LLM | llama3.2 (~2GB) |

---

## 📁 Auto-Generated Files

Libby creates these automatically — you never need to edit them:

| File | Purpose |
|---|---|
| `libby_config.json` | Saves your settings between sessions |
| `feedback_log.json` | Audit trail of every interaction and rating |
| `libby_eval_set.json` | Saved evaluation test cases |
| `libby_db/` | ChromaDB vector database — stores your indexed knowledge |

---

## 🗺️ Roadmap

- [ ] Tune evaluation scorer to reduce false negatives
- [ ] Package as `.exe` with PyInstaller — one-click launch, no Python required
- [ ] First-run setup wizard
- [ ] Deploy to offline hardware (Raspberry Pi / Mini PC)
- [ ] Solar power integration for true off-grid use
- [ ] OCR support for scanned PDFs (pytesseract)
- [ ] Persistent conversation history between sessions (sqlite3)
- [ ] Flashcard and quiz mode for personal knowledge base
- [ ] Cybersecurity layer — encryption (Fernet), bandit scanning
- [ ] Cloud / ERP connector (connection-optional architecture)
- [ ] Word document `.docx` support

---

## 📜 Version History

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
| **V13** | **Pillow splash screen + sidebar logo, xlsxwriter reports with charts, ruff linting** |

---

*Built with Python, ChromaDB, Ollama, sentence-transformers, Pillow and xlsxwriter*
*Designed for offline, air-gapped environments*
*Created by DanaBuilds — github.com/DanaBuilds/Offline-Libby*