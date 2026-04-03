# ─────────────────────────────────────────────
# ruff.toml — Libby V13 linting configuration
# Install:  pip install ruff
# Run:      ruff check libby_V13.py
# Fix:      ruff check libby_V13.py --fix
# ─────────────────────────────────────────────

# Target Python version
target-version = "py312"

# Max line length — relaxed slightly for long UI strings
line-length = 120

[lint]
# Rule sets to enable
select = [
    "E",    # pycodestyle errors
    "W",    # pycodestyle warnings
    "F",    # pyflakes (undefined names, unused imports)
    "I",    # isort (import ordering)
    "UP",   # pyupgrade (modernise syntax)
    "B",    # flake8-bugbear (common bugs)
    "SIM",  # flake8-simplify
]

# Rules to ignore — intentional style choices in Libby
ignore = [
    "E501",   # line too long — handled by line-length above
    "E741",   # ambiguous variable names (e.g. 'l', 'I') — used in UI loops
    "B006",   # mutable default arguments — not an issue in this codebase
    "SIM108", # ternary operator suggestions — readability preference
    "UP032",  # f-string suggestions — some kept as .format() for clarity
    "W291",   # trailing whitespace — auto-fixed
    "W293",   # whitespace before colon — auto-fixed
    "F841",   # local variable assigned but never used — some are intentional UI refs
]

# Files and folders to skip
exclude = [
    "__pycache__",
    "libby_db",
    ".git",
    "*.json",
]

[lint.isort]
# Keep stdlib, third-party and local imports grouped cleanly
known-third-party = [
    "chromadb",
    "sentence_transformers",
    "PIL",
    "pandas",
    "openpyxl",
    "xlsxwriter",
    "fitz",
    "watchdog",
    "requests",
    "customtkinter",
]