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

## Quick Start

```bash
pip install openpyxl
python quiz_app.py
```

Requires Python 3. The UI is built with the bundled Tkinter — no extra installation needed.

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

The app automatically recognizes common header variants and option markers (`A.` `A、` `A:` etc.). A sample file, `示例题库.xlsx`, is included for reference.

## Where Is My Data

All data is stored in the app directory, so you can copy it anywhere:

- `quiz_bank.db` — question bank and wrong-answer notebook
- `quiz_session.json` — practice progress

Delete the corresponding file to clear that data. Back them up before upgrading.

## About AI Explanations

AI features are entirely optional — everything works without configuration. Only when you click "AI Explanation" will the current question be sent to the model service you configured.

## Terms of Use

This project is intended for personal study and question-bank review. Do not use it in exams, competitions, or formal assessments where external tools or AI assistance are prohibited. Question content and AI-generated explanations may contain errors — always verify them yourself.

License: MIT
