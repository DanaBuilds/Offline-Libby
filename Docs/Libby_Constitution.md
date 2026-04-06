# Libby — Constitution
> Non-negotiable principles that govern every version of Libby, every session, every feature.
> Any AI agent, developer or contributor working on Libby must adhere to these principles before writing a single line of code.

---

## 1. Core Identity

- Libby is a **fully offline, air-gapped AI knowledge assistant**
- She runs entirely on the user's local machine — no cloud, no internet after initial setup
- She answers questions **only from the user's own documents** — never from outside knowledge
- Her name is **Libby** — this never changes regardless of deployment context
- Current version: **V13**

---

## 2. Offline & Air-Gap — Non-Negotiable

- Libby **never** makes outbound network calls during runtime
- Libby **never** sends user data, questions, answers or documents to any external server
- All AI inference runs locally via **Ollama**
- All vector search runs locally via **ChromaDB**
- All embeddings are generated locally via **sentence-transformers**
- The only permitted network activity is the one-time initial setup (pip installs, ollama pull)
- Libby must remain deployable in a **Faraday cage** with zero signal dependency

---

## 3. Design Language — Non-Negotiable

- **Primary theme:** Charcoal dark (`#1e1e1e`, `#252525`) with rose gold accents (`#c9956c`, `#8b5e3c`)
- **Secondary theme:** Warm linen light (`#f5f0eb`, `#ffffff`) with muted rose gold accents (`#9b5e3a`)
- **Font:** Segoe UI throughout — no other font family is permitted
- **Font sizes:** Title 26px bold, chat 11px, labels 9px, tiny 7–8px
- **UI framework:** tkinter + CustomTkinter — no web frameworks
- No gradients, no drop shadows, no neon effects
- Rose gold is Libby's signature colour — it must appear in every version in headers, accents, source citations and report branding
- All exported Excel reports must carry the charcoal and rose gold branding

---

## 4. Supported File Types

Libby reads and indexes the following file types only:
- `.txt` — plain text knowledge files
- `.xlsx` — Excel spreadsheets for knowledge and BI data
- `.pdf` — text-based PDF documents (scanned PDFs require OCR — not yet implemented)

No other file types are permitted without a deliberate version upgrade decision.

---

## 5. Architecture Principles

- **Single file** — Libby lives in one `.py` file until the codebase is feature-complete and stable
- **No hardcoded paths** — all paths use `SCRIPT_DIR` and `os.path` for portability
- **Auto path detection** — Libby works on any machine without editing the source code
- **Config persistence** — all user settings saved to `libby_config.json` between sessions
- **Section dividers** — all code sections separated by clear `# ─────` divider comments
- **Graceful fallbacks** — optional libraries (watchdog, Pillow) must fail silently with a print warning, never a crash
- **Threading** — all RAG queries and backend loading run in daemon threads, never on the main UI thread

---

## 6. RAG Pipeline Standards

- **Chunk size:** 1200 characters
- **Chunk overlap:** 150 characters
- **Cosine distance filter:** 1.4 maximum (chunks beyond this threshold are discarded)
- **Results returned:** top 5 chunks per query
- **Embedding model:** `all-MiniLM-L6-v2` via sentence-transformers
- **Context window passed to LLM:** last 4 conversation exchanges only
- **Conversation history stored:** last 20 exchanges per tab
- Libby **never guesses** — if the answer is not in the knowledge base she says so explicitly

---

## 7. Tab System

Libby has exactly three tabs — each with fully independent conversation state:

| Tab | Purpose | End User Facing |
|-----|---------|----------------|
| 📚 Knowledge Assistant | Document Q&A | Yes |
| 📊 Enterprise BI | Data, calculations, reports | Yes |
| 🧪 Evaluation | Accuracy testing | No — developer only |

- Each tab maintains its own conversation history
- Each tab has its own dot indicator (●) when it contains messages
- The Clear button only clears the active tab
- Tab order never changes

---

## 8. MCP — File Watching

- Libby monitors her knowledge folder live using `watchdog` (LibbyMCPHandler + LibbyMCP)
- Supported events: file added, modified, deleted, renamed
- Debounce delay: **2 seconds** — prevents duplicate re-index events
- MCP notifies the user in chat when a file change is detected
- MCP restarts automatically when the knowledge folder changes in Settings
- If watchdog is not installed Libby falls back gracefully — never crashes

---

## 9. Audit & Feedback

- Every Libby response receives a unique ID
- Every interaction is logged to `feedback_log.json` — timestamp, tab, question, answer, sources
- Every Libby bubble has 👍 👎 buttons
- Ratings update the matching log entry in place
- The audit log is never deleted automatically — it is the user's data

---

## 10. Evaluation System

- Test cases stored in `libby_eval_set.json`
- Scoring uses keyword matching — a result passes if at least half the expected keywords appear in Libby's answer
- The evaluation tab is a developer tool — it must never be marketed as an end-user feature
- Evaluation runs silently through the full RAG pipeline — it does not use shortcuts

---

## 11. Report Engine

- Reports are generated using **xlsxwriter** — not openpyxl (openpyxl is retained for reading only)
- Every report includes: company header, subtitle with timestamp, sorted data rows, totals row, bar chart, confidentiality footer
- Report branding must always use charcoal (`#252525`) and rose gold (`#c9956c`, `#8b5e3c`)
- The confidentiality footer always reads: `Confidential • [Company Name] • [Date] • Air-gapped`
- Reports are triggered by the phrase **"generate report"** in the Enterprise BI tab

---

## 12. Settings

Users may configure the following — nothing else is user-configurable without a version decision:

| Setting | Options |
|---------|---------|
| Knowledge folder | Any valid local folder path |
| Theme | Dark (charcoal) or Light (warm linen) |
| Company name | Any string, defaults to "Libby Intelligence" |
| AI model | Any installed Ollama LLM |

- Settings persist in `libby_config.json`
- Changing the knowledge folder triggers a full re-index
- Changing the model takes effect immediately — no restart required

---

## 13. Naming & Versioning

- File naming convention: `libby_V[N].py` — always uppercase V
- Version increments with every meaningful feature addition or fix session
- The window title always reflects the current version: `Libby V[N] — Universal Knowledge System`
- Welcome messages always reflect the current version
- GitHub repository: `github.com/DanaBuilds/Offline-Libby`

---

## 14. Auto-Generated Files

These files are created automatically and must never be deleted by Libby's own code:

| File | Purpose |
|------|---------|
| `libby_config.json` | User settings |
| `feedback_log.json` | Interaction audit log |
| `libby_eval_set.json` | Evaluation test cases |
| `libby_db/` | ChromaDB vector database |
| `ruff.toml` | Linting configuration |

---

## 15. Roadmap Principles

Features are added in this priority order:
1. Stability and bug fixes always before new features
2. Core RAG quality improvements before UI enhancements
3. Security features before cloud/network features
4. Deployment readiness (PyInstaller, first-run wizard) before platform expansion

---

*This constitution was established at V13.*
*Created by DanaBuilds — github.com/DanaBuilds/Offline-Libby*