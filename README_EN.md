# Offline Quiz App

**English** | [中文](README.md)

A simple, local-first quiz app: import your Excel question bank and start practicing right away.

No internet connection, no sign-up — all data stays on your own computer.

## What It Does

- **Import & go** — supports `.xlsx` question banks; each Sheet becomes a separate test paper, imported once and saved permanently
- **Three question types** — single-choice, multiple-choice, and true/false, with instant scoring and explanations
- **Flexible ordering** — random shuffle / original order / grouped by type, switchable anytime
- **Wrong-answer notebook** — questions you get wrong are collected automatically and removed once answered correctly
- **Auto-saved progress** — close the app and reopen it; you'll continue right where you left off
- **Smart filtering** — filters out questions that only require memorizing keywords, generating a condensed review bank
- **AI explanations (optional)** — connect to OpenAI-compatible APIs, Ollama, LM Studio, and more to generate key points and common pitfalls with one click

## Installation

Requires **Python 3.9 or newer** (3.11 recommended). The UI is built with the bundled Tkinter — no extra GUI framework needed.

Using a virtual environment is recommended to avoid polluting your system Python. Pick either option below.

### Option 1: venv (recommended)

Windows (PowerShell):

```powershell
# 1. Go to the project directory
cd D:\path\to\quiz_app

# 2. Create a virtual environment
python -m venv .venv

# 3. Activate it
.venv\Scripts\Activate.ps1
# If script execution is blocked, use instead: .venv\Scripts\activate.bat

# 4. Install dependencies
pip install -r requirements.txt
```

macOS / Linux:

```bash
cd /path/to/quiz_app
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Option 2: Conda

```bash
# 1. Create and activate an environment
conda create -n quiz python=3.11
conda activate quiz

# 2. Install dependencies
cd /path/to/quiz_app
pip install -r requirements.txt
```

### Minimal install

If you'd rather skip virtual environments, installing the single third-party dependency is enough:

```bash
pip install openpyxl
```

## Running

Activate the environment first (venv: `.venv\Scripts\Activate.ps1`; conda: `conda activate quiz`), then:

```bash
python main.py
# equivalent: python -m quiz_app
```

> **Note:** If you see `No module named 'tkinter'`, your Python is missing the Tk component:
> - Windows: re-run the official installer and check **"tcl/tk and IDLE"**
> - Ubuntu/Debian: `sudo apt install python3-tk`
> - macOS: prefer the python.org installer (the Homebrew build may lack Tk)

## Question Bank Format

The first row of your Excel file is the header. Only the **Question** column is required; everything else is optional:

| Column | Required | Description |
| --- | --- | --- |
| 题目 (Question) | ✅ | Question stem |
| 题型 (Type) | No | single / multiple / true-false; defaults to single |
| 选项 (Options) | Recommended for single/multiple | All options in one cell, or split into columns like 选项A, 选项B… |
| 答案 (Answer) | No | e.g. `A`, `A,B`, `对/错` (true/false), or the answer text itself |
| 解析 (Explanation) | No | Answer explanation |
| 难度 (Difficulty) | No | 简单 / 中等 / 困难 (easy / medium / hard) |

The app automatically recognizes common header variants and option markers (`A.` `A、` `A:` etc.). See `samples/示例题库.xlsx` for reference, or regenerate it with `python scripts/make_sample.py`.

## Project Layout

```text
quiz_app/
├── main.py          # entry point (python main.py)
├── quiz_app/        # application package
│   ├── app.py                  # main window and quiz flow
│   ├── ai_explanation.py       # AI explanations (OpenAI-compatible / Ollama)
│   ├── dialogs.py              # dialogs
│   ├── settings.py             # appearance and answering settings
│   ├── shortcuts.py            # keyboard shortcuts
│   ├── memorization_filter.py  # smart filtering rules
│   └── paths.py                # unified paths (repo root / data / samples)
├── scripts/         # command-line scripts
├── tests/           # tests (python tests/test_core.py)
├── samples/         # sample question bank
└── data/            # runtime data, not committed
```

## Where Is My Data

All runtime data lives in `data/`, so you can copy it anywhere:

- `data/quiz_bank.db` — question bank and wrong-answer notebook
- `data/quiz_session.json` — practice progress
- `data/*_settings.json` — appearance, shortcuts, and AI endpoint settings

Delete the corresponding file to clear that data. Back up the whole `data/` folder before upgrading.

## Tests

```bash
python tests/test_core.py        # bank / session / scoring
python tests/test_import.py      # Excel import parsing
python tests/test_gui_smoke.py   # GUI smoke test (opens a window for ~2 seconds)
python -m unittest discover -s tests -p "test_*.py"
```

Tests write to a temporary directory by default, so your real data in `data/` is never touched.

## About AI Explanations

AI features are entirely optional — everything works without configuration. Only when you click "AI Explanation" will the current question be sent to the model service you configured.

## Terms of Use

This project is intended for personal study and question-bank review. Do not use it in exams, competitions, or formal assessments where external tools or AI assistance are prohibited. Question content and AI-generated explanations may contain errors — always verify them yourself.

## License

Released under the [MIT License](LICENSE) — free to use, modify, and distribute.
