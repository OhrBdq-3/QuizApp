# -*- coding: utf-8 -*-
"""统一路径管理。

目录约定（APP_ROOT = 仓库根目录）：
    APP_ROOT/
        quiz_app/    源码包
        scripts/     命令行脚本
        tests/       测试
        samples/     示例题库
        data/        运行时数据（题库 / 会话 / 各项设置，不入库）
        outputs/     脚本生成的产物，不入库

打包成 exe 后 APP_ROOT 取 exe 所在目录，数据目录仍为其下的 data/。
"""
from __future__ import annotations

import os
import sys

if getattr(sys, "frozen", False):          # PyInstaller 等打包环境
    APP_ROOT = os.path.dirname(sys.executable)
else:
    # 本文件位于 <APP_ROOT>/quiz_app/paths.py
    APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DATA_DIR = os.path.join(APP_ROOT, "data")
SAMPLES_DIR = os.path.join(APP_ROOT, "samples")
OUTPUTS_DIR = os.path.join(APP_ROOT, "outputs")

# 历史上这些数据文件曾与源码同放在根目录，重构后统一收进 data/
_LEGACY_FILES = (
    "quiz_bank.db",
    "quiz_session.json",
    "app_settings.json",
    "ai_settings.json",
    "shortcut_settings.json",
    "示例题库.progress.json",
)


def ensure_dirs() -> None:
    """确保运行时目录存在。"""
    os.makedirs(DATA_DIR, exist_ok=True)


def migrate_legacy_data() -> list[str]:
    """把仍留在根目录的历史数据文件搬进 data/，返回被迁移的文件名列表。"""
    ensure_dirs()
    moved = []
    for name in _LEGACY_FILES:
        old = os.path.join(APP_ROOT, name)
        new = os.path.join(DATA_DIR, name)
        if os.path.isfile(old) and not os.path.exists(new):
            try:
                os.replace(old, new)
                moved.append(name)
            except OSError:
                pass
    return moved


def data_path(name: str) -> str:
    """data/ 目录下某个文件的完整路径。"""
    return os.path.join(DATA_DIR, name)


ensure_dirs()
migrate_legacy_data()
