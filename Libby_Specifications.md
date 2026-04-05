# Libby — Specifications
> Describes WHAT Libby does and WHY — not HOW it is implemented.
> Each specification defines user needs, functional requirements, and acceptance criteria.
> All specifications must comply with constitution.md at all times.
> Living document — updated each version as features are added or changed.

**Current version:** V13
**Last updated:** April 2026

---

## SPEC-001 — RAG Knowledge Assistant

### Overview
Users need to ask questions in plain English and receive accurate, sourced answers drawn exclusively from their own document library — without internet access or technical knowledge.

### User Stories
- As a user, I want to ask a question and receive a clear, complete answer so that I can find information quickly without searching through files manually
- As a user, I want to know which file the answer came from so that I can verify or read more detail
- As a user, I want Libby to tell me honestly when she doesn't know something so that I never act on a wrong answer
- As a user, I want to ask follow-up questions naturally so that I can explore a topic in conversation without repeating context

### Functional Requirements
- FR-001: Libby MUST answer questions using only content from the loaded knowledge base
- FR-002: Libby MUST cite the source filename for every answer
- FR-003: Libby MUST respond with "Not found in knowledge base" when no relevant content exists
- FR-004: Libby MUST detect follow-up questions and enrich the search query with prior context
- FR-005: Libby MUST retain the last 20 conversation exchanges per tab
- FR-006: Libby MUST pass only the last 4 exchanges to the LLM to manage context window

### Acceptance Criteria
- AC-001: A question with a clear answer in the knowledge base returns a correct, sourced response
- AC-002: A question with no relevant content returns a clear "not found" message — never a hallucinated answer
- AC-003: A follow-up like "tell me more" returns content relevant to the previous topic without re-stating the question
- AC-004: Source citation shows the correct filename for the answer

### Out of Scope
- Web search or external knowledge
- Answering from memory or training data
- Multi-language support (English only for now)

---

## SPEC-002 — Enterprise BI & Data Calculations

### Overview
Users with Excel-based business data need to ask calculation questions and receive precise numeric answers — totals, averages, maximums, minimums and counts — without writing formulas or opening spreadsheets.

### User Stories
- As a business user, I want to ask "what is the total revenue?" and get the correct calculated answer so that I can get data insights without opening Excel
- As a business user, I want to ask "who has the highest salary?" and get the name and value so that I can identify top performers or outliers instantly
- As a business user, I want answers to be concise and numeric so that I can use them in decisions without reading paragraphs

### Functional Requirements
- FR-001: Libby MUST perform sum, average, maximum, minimum and count calculations on numeric columns
- FR-002: Libby MUST format currency columns with $ and two decimal places
- FR-003: Libby MUST identify the relevant Excel file and sheet automatically from the question
- FR-004: Libby MUST return numeric answers in 3 lines or fewer
- FR-005: Libby MUST show the calculation formula when adding values (e.g. $6,750 + $4,800 = $11,550)
- FR-006: Libby MUST support follow-up lookups — when the user references "that" or "it", look up the full matching row

### Acceptance Criteria
- AC-001: "What is the total salary?" returns the correct summed value with $ formatting
- AC-002: "Who has the highest revenue?" returns the correct row identifier and value
- AC-003: A vague follow-up like "show me that record" returns the full matching row from the previous answer
- AC-004: Answers never exceed 3 lines

### Out of Scope
- Writing back to Excel files
- Creating pivot tables
- Filtering by date ranges (future feature)

---

## SPEC-003 — Excel Report Generation

### Overview
Users need to export professionally branded, formatted Excel reports from their data with a single conversational command — without knowing Excel formatting or having to manually build reports.

### User Stories
- As a business user, I want to type "generate report showing sales by region" and receive a formatted Excel file so that I can share professional reports without manual work
- As a business user, I want reports to be sorted and grouped automatically so that the most important data is immediately visible
- As a business user, I want reports to include a visual chart so that the data tells a story at a glance
- As a business user, I want reports to carry my company branding so that they look professional when shared

### Functional Requirements
- FR-001: Report generation MUST be triggered by the phrase "generate report" in the Enterprise BI tab
- FR-002: Libby MUST present a save dialog for the user to choose the output location
- FR-003: Reports MUST include: company header, date/time subtitle, column headers, data rows, totals row, bar chart and confidentiality footer
- FR-004: Reports MUST apply alternating row colours, currency formatting and sorted data
- FR-005: Reports MUST be branded in charcoal and rose gold — no other colour scheme permitted
- FR-006: Reports MUST include an auto-generated bar chart from the first label and numeric columns
- FR-007: Libby MUST confirm successful generation with a row count and sort direction summary

### Acceptance Criteria
- AC-001: "generate report showing sales by region highest to lowest" produces a correctly sorted, grouped Excel file
- AC-002: The report opens in Excel with correct branding, chart and totals row
- AC-003: The confidentiality footer reads: Confidential • [Company Name] • [Date] • Air-gapped
- AC-004: Cancelling the save dialog shows "Report cancelled" — no file is created

### Out of Scope
- PDF report export (future feature)
- Word document export (future feature)
- Scheduled or automated report generation

---

## SPEC-004 — MCP Live File Monitoring

### Overview
Users need their knowledge base to update automatically when files change — without restarting Libby — so that new information is immediately available without manual intervention.

### User Stories
- As a user, I want to drop a new file into my knowledge folder and have Libby index it automatically so that I don't have to restart the app
- As a user, I want to be notified in the chat when a file change is detected so that I know the knowledge base is being updated
- As a user on a deployment device, I want to add files via USB and have Libby pick them up automatically so that the air-gapped workflow is seamless

### Functional Requirements
- FR-001: Libby MUST monitor the knowledge folder continuously for file additions, modifications, deletions and renames
- FR-002: Libby MUST trigger a re-index automatically when a supported file type changes
- FR-003: Libby MUST display a notification message in the chat confirming the detected change and re-index
- FR-004: Libby MUST apply a 2 second debounce to prevent duplicate re-index events from a single save operation
- FR-005: Libby MUST restart the file watcher automatically when the knowledge folder changes in Settings
- FR-006: Libby MUST stop the file watcher cleanly when the app closes
- FR-007: If watchdog is not installed Libby MUST fall back gracefully — never crash

### Acceptance Criteria
- AC-001: Adding a new .txt file to the knowledge folder triggers a re-index and chat notification within 3 seconds
- AC-002: Deleting a file triggers a re-index
- AC-003: Rapid repeated saves to the same file trigger only one re-index
- AC-004: Closing Libby with the watcher running produces no errors

### Out of Scope
- Monitoring multiple folders simultaneously (one folder at a time)
- Network drive monitoring

---

## SPEC-005 — Audit Log & Feedback System

### Overview
Users and administrators need a verifiable record of every interaction with Libby — including user ratings — so that accuracy can be tracked, reviewed and improved over time.

### User Stories
- As a user, I want to rate Libby's answers with thumbs up or down so that I can flag incorrect or unhelpful responses
- As an administrator, I want a log of every question and answer so that I can audit Libby's accuracy and behaviour
- As a developer, I want ratings tied to specific log entries so that I can identify which answers need improvement

### Functional Requirements
- FR-001: Every Libby response MUST be assigned a unique ID at generation time
- FR-002: Every interaction MUST be logged to feedback_log.json with: ID, timestamp, tab, question, answer, sources
- FR-003: Every Libby chat bubble MUST display 👍 and 👎 buttons
- FR-004: A rating MUST update the matching log entry in place — not create a duplicate
- FR-005: Rated buttons MUST provide immediate visual feedback (rose gold for positive, red for negative)
- FR-006: The audit log MUST never be deleted or truncated by Libby automatically

### Acceptance Criteria
- AC-001: Every answer has a unique ID that appears in feedback_log.json
- AC-002: Clicking 👍 updates the correct log entry with rating "positive"
- AC-003: The log file persists between sessions and accumulates all interactions
- AC-004: Rating buttons change colour immediately on click

### Out of Scope
- Cloud sync of audit logs
- Exporting audit logs to Excel (future feature)
- Admin dashboard UI for log review (future feature)

---

## SPEC-006 — Evaluation System

### Overview
Developers need a way to test Libby's RAG accuracy against a defined set of questions and expected answers — so that regressions can be caught before they reach users, and improvements can be measured objectively.

### User Stories
- As a developer, I want to add question and expected answer pairs so that I can define what correct behaviour looks like
- As a developer, I want to run all tests silently and see a score so that I can measure Libby's accuracy at a glance
- As a developer, I want to see which specific questions passed or failed so that I can diagnose problems precisely

### Functional Requirements
- FR-001: The evaluation tab MUST allow adding question + expected answer pairs
- FR-002: Test cases MUST persist between sessions in libby_eval_set.json
- FR-003: Running evaluation MUST send each question through the full RAG pipeline silently
- FR-004: Scoring MUST use keyword matching — pass if at least half the expected keywords appear in the answer
- FR-005: Results MUST show per-question ✅ ❌ with question preview and answer preview
- FR-006: Final score MUST be colour coded — green 70%+, amber 40–69%, red below 40%
- FR-007: The evaluation tab MUST NOT be presented as an end-user feature

### Acceptance Criteria
- AC-001: Adding a test case and running evaluation produces a ✅ or ❌ result for that question
- AC-002: Test cases survive closing and reopening Libby
- AC-003: A score of 0% shows in red, 50% in amber, 80% in green
- AC-004: Evaluation runs do not affect the Knowledge Assistant conversation history

### Out of Scope
- Semantic similarity scoring (keyword matching only for now)
- Automated scheduled evaluation runs
- Evaluation of the Enterprise BI tab

---

## SPEC-007 — Settings & Configuration

### Overview
Users need to configure Libby for their environment — folder location, theme, company name and AI model — without touching the source code, so that Libby works on any machine for any organisation.

### User Stories
- As a user, I want to point Libby at any folder on my computer so that I can use any knowledge base I choose
- As a user, I want to choose between dark and light themes so that Libby fits my working environment
- As a user, I want to enter my company name so that it appears in the header and exported reports
- As a user, I want to switch AI models without restarting so that I can test different LLMs easily

### Functional Requirements
- FR-001: Settings MUST be accessible from a panel opened by the ⚙ Settings button
- FR-002: All settings MUST persist to libby_config.json between sessions
- FR-003: Changing the knowledge folder MUST trigger a full re-index automatically
- FR-004: Changing the theme MUST apply immediately without restarting
- FR-005: Changing the AI model MUST take effect on the next query — no restart required
- FR-006: The knowledge folder path MUST be validated — Libby MUST reject non-existent paths

### Acceptance Criteria
- AC-001: Closing and reopening Libby restores all previous settings
- AC-002: Changing the folder path shows "Reloading..." and re-indexes the new folder
- AC-003: Switching from dark to light theme changes all UI colours immediately
- AC-004: Entering a non-existent folder path shows an error and does not save

### Out of Scope
- Per-tab settings
- Password protection of settings
- Remote configuration

---

## SPEC-008 — Splash Screen & Visual Identity (V13)

### Overview
Users need a polished, branded startup experience that communicates Libby's identity and offline status immediately — so that the app feels professional and purpose-built from the first moment.

### User Stories
- As a user, I want to see a branded splash screen on startup so that the app feels professional and intentional
- As a user, I want the sidebar to have a visual identity element so that Libby feels like a named product, not a generic tool

### Functional Requirements
- FR-001: A splash screen MUST display for 2.5 seconds on startup
- FR-002: The splash screen MUST display Libby's name, version, and "Offline • Air-gapped" status
- FR-003: The splash screen MUST use the charcoal and rose gold colour scheme
- FR-004: The splash screen MUST close automatically and return full input focus to the main window
- FR-005: The sidebar MUST display a rose gold logo icon above the Knowledge Base label
- FR-006: If Pillow is not installed both features MUST degrade gracefully — never crash

### Acceptance Criteria
- AC-001: Splash screen appears on launch, displays for ~2.5 seconds and closes cleanly
- AC-002: The input box is fully interactive immediately after the splash closes
- AC-003: The sidebar logo is visible in both dark and light themes
- AC-004: Running Libby without Pillow installed produces no errors — splash and logo are simply skipped

### Out of Scope
- Animated splash screens
- User-customisable splash screen content
- Logo file loading from disk (logo is generated programmatically)

---

## Upcoming Specifications (Planned)

These features are on the roadmap and will receive full specifications before implementation begins:

| Spec ID | Feature | Priority |
|---------|---------|---------|
| SPEC-009 | OCR for scanned PDFs | High |
| SPEC-010 | Persistent conversation history (sqlite3) | High |
| SPEC-011 | PyInstaller packaging (.exe) | High |
| SPEC-012 | First-run setup wizard | Medium |
| SPEC-013 | Encryption & cybersecurity layer (Fernet) | Medium |
| SPEC-014 | Flashcard & quiz mode | Medium |
| SPEC-015 | Raspberry Pi / Mini PC deployment | Medium |
| SPEC-016 | Word document (.docx) support | Low |
| SPEC-017 | Cloud / ERP connector | Low |

---

*Specifications established at V13.*
*All specs must remain aligned with constitution.md.*
*Created by DanaBuilds — github.com/DanaBuilds/Offline-Libby*