# -*- coding: utf-8 -*-
"""离线刷题 App。

用法：
    python main.py          # 仓库根目录下的入口脚本
    python -m quiz_app      # 等价写法
"""
from .app import (
    QuizApp,
    BankDB,
    Question,
    load_excel,
    save_session,
    load_session,
    main,
)
from .paths import APP_ROOT, DATA_DIR, SAMPLES_DIR

__all__ = [
    "QuizApp",
    "BankDB",
    "Question",
    "load_excel",
    "save_session",
    "load_session",
    "main",
    "APP_ROOT",
    "DATA_DIR",
    "SAMPLES_DIR",
]
