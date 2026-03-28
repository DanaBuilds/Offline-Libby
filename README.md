# 📚 Libby V7 — Universal Offline AI Assistant

Libby is a fully offline, air-gapped AI assistant that answers questions from your personal knowledge base and business data. She runs entirely on your machine — no internet, no cloud, no data leaving your computer.

---

## ✨ What's New in V7

- **Universal** — one app, two modes in separate tabs
- **Settings panel** — change your knowledge folder, theme and company name without touching code
- **Tab system** — Knowledge Assistant and Enterprise BI each have their own independent conversation
- **Dot indicator** — shows when a tab has conversation history
- **Smart report generation** — type "generate report..." to export formatted Excel files
- **Folder switching** — point Libby at any folder on your computer, she rebuilds automatically
- **No hardcoded paths** — works on any computer without editing the code

---

## 🖥️ System Requirements

- Windows 10 or 11 (64-bit)
- At least 8GB RAM (16GB recommended)
- At least 10GB free disk space
- Internet connection for initial setup only

---

## 🗂️ Folder Structure

Create a folder anywhere on your computer for your knowledge base. Libby will find all supported files automatically no matter how deep the subfolders go.

Example structure:
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

Libby auto-creates a `libby_db` folder inside your knowledge folder to store her database. You never need to touch it.

---

## ⚙️ Step 1 — Install Python

1. Go to `https://www.python.org/downloads/`
2. Download the latest **Python 3.x** version for Windows (64-bit)
3. Run the installer
4. ⚠️ **Important:** Check **"Add Python to PATH"** before clicking Install
5. Click **Install Now**

Verify it worked — open Command Prompt and type:
```
python --version
```
You should see something like `Python 3.14.0`

---

## 📦 Step 2 — Install Required Libraries

Open Command Prompt and run:
```
pip install chromadb sentence-transformers openpyxl pymupdf requests pandas
```

This may take a few minutes. Wait for it to finish before moving on.

---

## 🤖 Step 3 — Install Ollama

Ollama runs the AI brain locally on your machine.

1. Go to `https://ollama.com/download`
2. Click **Download for Windows**
3. Run the installer like any normal program

Verify it worked:
```
ollama --version
```

---

## 🧠 Step 4 — Download the AI Model

Libby uses `phi3:mini` — a fast, efficient model that works well on most machines.

In Command Prompt:
```
ollama pull phi3:mini
```

This downloads approximately 2GB. This is the **last time you need internet** for Libby to work.

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
- Use clear descriptive filenames
- Organise files into subfolders by topic
- Excel files work best with clear column headers in row 1
- PDFs must be text-based (not scanned images)

---

## 🚀 Step 6 — Run Libby

Run from Command Prompt:
```
python libby_V7.py
```

On first launch Libby will index all your files. This may take a minute depending on how many files you have. Subsequent launches are much faster.

---

## ⚙️ Step 7 — Configure Your Settings

Click **⚙ Settings** in the top right corner to configure:

| Setting | What it does |
|---|---|
| **Knowledge Folder** | Point Libby at your files — click Browse to select any folder |
| **Theme** | Switch between Dark Mode and Light Mode |
| **Company Name** | Personalise the header and exported reports |

Click **Save & Apply** — changes take effect immediately, no restart needed.

When you change the knowledge folder Libby automatically clears her database and rebuilds from the new location. Only files from the selected folder will appear.

---

## 💬 How to Use Libby

### 📚 Knowledge Assistant Tab

Ask questions in plain English from your loaded documents:

```
How do I treat a second degree burn?
What foods have the longest shelf life?
How long should I boil water to purify it?
```

Libby answers using only what is in your knowledge base — never guessing or making things up.

### 📊 Enterprise BI Tab

Ask data questions from your Excel files:

```
What is the total salary budget?
Which product has the highest profit margin?
Which customers are on hold?
What is the average revenue per sales rep?
```

Libby performs real calculations using your actual data.

### 📊 Generating Reports

In the Enterprise BI tab, start your message with **"generate report"**:

```
generate report showing sales by region highest to lowest
generate report of inventory items lowest to highest stock
generate report of active customers by outstanding balance
generate report of pending purchase orders by supplier
generate report showing employees by salary highest to lowest
```

A save dialog opens — choose where to save your Excel report. Libby builds a formatted report with a company header, sorted and grouped data, a totals row and a confidentiality footer.

---

## 🗂️ Tab System

Libby has two independent conversation tabs:

| Tab | Purpose | Indicator |
|---|---|---|
| 📚 Knowledge Assistant | Document Q&A | ● dot appears when tab has messages |
| 📊 Enterprise BI | Data, calculations, reports | ● dot appears when tab has messages |

- Each tab keeps its own separate conversation history
- Switching tabs instantly shows that tab's conversation
- The **Clear** button only clears the active tab
- History resets when you close Libby

---

## 🔒 Privacy & Air-Gap

Libby is designed to work **completely offline:**

- ✅ No data sent to the internet
- ✅ No cloud APIs used
- ✅ All knowledge stays on your machine
- ✅ Ollama and ChromaDB run entirely locally
- ✅ Generated reports never leave your computer unless you share them

After the initial setup Libby never needs an internet connection.

---

## 🛠️ Troubleshooting

**Libby shows "0 chunks" or files are missing:**
- Make sure your files are inside the selected knowledge folder
- Delete the `libby_db` folder inside your knowledge folder and restart
- Check that your files are `.txt`, `.xlsx` or `.pdf` format

**"Ollama not running" error:**
- Open a new Command Prompt window and type: `ollama serve`
- Leave that window open and restart Libby

**Excel file not loading:**
- Open the file in LibreOffice Calc or Excel and re-save as `.xlsx`
- Make sure the file is not open in another program when Libby starts
- Check that row 1 contains column headers

**Settings not saving:**
- Make sure the selected folder actually exists on your computer
- Check that you have write permission to the Libby folder

**File saved as `.py.txt` instead of `.py`:**
- In Notepad: File → Save As → change Save as type to **All Files (*.*)**
- Name the file `libby_V7.py` exactly

---

## 📋 Requirements Summary

| Requirement | Version |
|---|---|
| Python | 3.10 or higher |
| chromadb | latest |
| sentence-transformers | latest |
| openpyxl | latest |
| pymupdf | latest |
| requests | latest |
| pandas | latest |
| Ollama | latest |
| AI Model | phi3:mini (~2GB) |

---

## 📁 Auto-Generated Files

Libby creates these automatically — you never need to edit them:

| File | Purpose |
|---|---|
| `libby_config.json` | Saves your settings (folder path, theme, company name) |
| `libby_db/` | ChromaDB vector database — stores your indexed knowledge |

---

## 🗺️ Roadmap

- [ ] Package as `.exe` for one-click launch — no Python required
- [ ] Deploy to dedicated offline hardware (Raspberry Pi / Mini PC)
- [ ] Solar power integration for true off-grid use
- [ ] Word document `.docx` support
- [ ] Scanned PDF support via OCR
- [ ] Conversation history saved between sessions

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
| **V7** | **Universal app, tab system, settings panel, folder switching** |

---

*Built with Python, ChromaDB, Ollama and phi3:mini*
*Designed for offline, air-gapped environments*
*Created by DanaBuilds*