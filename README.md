# 📚 Libby — Offline AI Knowledge Assistant

Libby is a fully offline, air-gapped AI assistant that answers questions from your personal knowledge base.
She reads `.txt`, `.xlsx`, and `.pdf` files and uses a local AI brain to give clear, accurate answers — no internet required.

---

## 🖥️ What You Need

Before setting up Libby, make sure your computer has:
- Windows 10 or 11 (64-bit)
- At least 8GB RAM (16GB or more recommended)
- At least 10GB of free disk space
- An internet connection for the initial setup only

---

## 🗂️ Folder Structure

Create a folder called `RAG Test` on your `C:` drive:

```
C:\RAG Test\
  ├── libby_ui.py          ← Libby V1 (friendly, conversational)
  ├── libby_ui_V2.py       ← Libby V2 (fast, strict, bullet points)
  ├── libby_ui_V3.py       ← Libby V3 (latest - recommended)
  ├── rag_test.py          ← Command line version
  ├── Farming\             ← Example knowledge folder
  │     ├── Fruit\
  │     ├── Vegetables\
  │     └── planting_guide.xlsx
  └── First Aid\           ← Example knowledge folder
```

You can rename or reorganize folders however you like.
Libby will automatically find all `.txt`, `.xlsx`, and `.pdf` files no matter how deep the folders go.

---

## ⚙️ Step 1 — Install Python

1. Go to `https://www.python.org/downloads/`
2. Download the latest **Python 3.x** version for Windows
3. Run the installer
4. ⚠️ **Important:** Check the box that says **"Add Python to PATH"** before clicking Install
5. Click **Install Now**

Verify it worked — open Command Prompt and type:
```
python --version
```
You should see something like `Python 3.14.0`

---

## 📦 Step 2 — Install Required Libraries

Open Command Prompt, navigate to your RAG Test folder:
```
cd "C:\RAG Test"
```

Then install all required libraries:
```
pip install chromadb sentence-transformers openpyxl pymupdf requests
```

This may take a few minutes. Wait for it to finish.

---

## 🤖 Step 3 — Install Ollama

Ollama is the engine that runs Libby's AI brain locally on your machine.

1. Go to `https://ollama.com/download`
2. Click **Download for Windows**
3. Run the installer like any normal program

Verify it worked — open a new Command Prompt and type:
```
ollama --version
```

---

## 🧠 Step 4 — Download the AI Model

Libby V3 uses `phi3:mini` — a fast, efficient model that runs well on most machines.

In Command Prompt type:
```
ollama pull phi3:mini
```

This downloads about 2GB. This is the **last time you need internet** for Libby to work.

To test it works:
```
ollama run phi3:mini
```
Type `hello` and it should respond. Type `/bye` to exit.

---

## 📄 Step 5 — Add Your Knowledge Files

Place your knowledge files inside `C:\RAG Test` in any folder structure you like.

Supported file types:
- 📄 `.txt` — plain text files
- 📊 `.xlsx` — Excel spreadsheets (save from LibreOffice or Excel)
- 📕 `.pdf` — text-based PDF documents (not scanned images)

**Tip:** LibreOffice is a free alternative to Microsoft Office that works perfectly for creating `.xlsx` and `.pdf` files.
Download it at `https://www.libreoffice.org/download/libreoffice/` — choose the **Still** version.

---

## 🚀 Step 6 — Run Libby

You need two Command Prompt windows open at the same time:

**Window 1 — Start Ollama:**
```
ollama serve
```
Leave this window open. It runs quietly in the background.

**Window 2 — Start Libby:**
```
cd "C:\RAG Test"
python libby_ui_V3.py
```

Libby will load, index your files and open her interface.
The first run takes longer as she indexes everything — subsequent runs are much faster.

---

## 💬 How to Use Libby

- Type your question in the box at the bottom
- Press **Enter** or click **QUERY**
- Libby will search your knowledge base and answer using only your files
- The source file is shown below each answer
- Use **🗑 Clear** to start a fresh conversation
- Toggle between **Dark Mode** and **Light Mode** using the button in the top right

---

## 🔒 Privacy & Air-Gap

Libby is designed to work **completely offline**:
- No data is sent to the internet
- No cloud APIs are used
- All knowledge stays on your machine
- Ollama and ChromaDB run entirely locally

After the initial setup, Libby never needs an internet connection.

---

## 🗃️ Libby Versions

| File | Version | Best For |
|---|---|---|
| `libby_ui.py` | V1 | Friendly conversational answers |
| `libby_ui_V2.py` | V2 | Fast strict bullet point answers |
| `libby_ui_V3.py` | V3 | Latest — recommended for most use |

---

## 🛠️ Troubleshooting

**"0 chunks loaded" on startup:**
- Delete the `rag_db` folder inside `C:\RAG Test` and restart Libby
- Make sure your files are inside `C:\RAG Test` or its subfolders

**"Ollama not running" error:**
- Make sure `ollama serve` is running in a separate CMD window

**Excel file not loading:**
- Open the file in LibreOffice Calc and re-save as `.xlsx` format
- Make sure it's not open in another program when Libby starts

**File saved as `.py.txt` instead of `.py`:**
- In Notepad, use **File → Save As → All Files (*.*)** when saving Python files
- Or rename using CMD: `ren "filename.py.txt" "filename.py"`

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
| Ollama | latest |
| AI Model | phi3:mini |

---

## 🗺️ Roadmap

- [ ] Package as `.exe` for one-click launch
- [ ] Deploy to dedicated offline hardware (Raspberry Pi / Mini PC)
- [ ] Solar power integration for true off-grid use
- [ ] Support for more file types (Word docs, CSV)

---

*Built with Python, ChromaDB, Ollama, and phi3:mini*
*Designed for offline, air-gapped environments*