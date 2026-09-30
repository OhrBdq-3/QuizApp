# -*- coding: utf-8 -*-
"""
离线刷题 App
============
功能：
  - 导入 Excel 题库（.xlsx / .xls），题库长期保存在本地 SQLite 数据库（quiz_bank.db）
  - 支持单选题 / 多选题 / 判断题
  - 逐题作答：单选/判断点选项即判分；多选需点「提交本题」
  - 每题判分后直接在界面上显示对错与正确答案（不弹窗）：
    答对约 0.7 秒后自动进入下一题，答错停留本题方便看解析
  - 可选择题目顺序：随机乱序 / 原顺序 / 题型分组（单选→多选→判断，组内原顺序）
  - 右侧题目列表，点击任意题目直接跳转
  - 错题自动收集（长期错题本），可单独重做错题
  - 自动保存刷题进度，关闭 App 后重新打开可从上次的题目继续

长期数据库 / 会话文件（与脚本或 exe 同目录）：
  quiz_bank.db      题库 + 错题本（SQLite）
  quiz_session.json 上次刷题位置（题库、顺序、进度、已答记录）

Excel 题库格式（每个 sheet 是一张试卷/题库）：
  表头行（第 1 行）必须包含以下列（顺序无关，缺省列用默认值）：
    题型       单选 / 多选 / 判断（也接受 单选题/多选题/判断题 等写法）
    题目       题干
    选项       A. xxx B. xxx C. xxx D. xxx
               （兼容 无空格、换行分隔、A、/A:/A) 等写法；
                 也支持把选项拆成多列：选项A 选项B 选项C… / 选项1 选项2…）
    答案       正确答案，多个答案用逗号/空格分隔，如 "A" 或 "A,B" 或 "AB"
               （判断题可写 对/错、√/×；也可直接写选项文本）
    解析       答案解析（可选）
    难度       简单 / 中等 / 困难（可选，默认 中等）

运行：
  python quiz_app.py
"""
from __future__ import annotations

import json
from .ai_explanation import AIExplanationMixin
from .shortcuts import ShortcutMixin
from .dialogs import AppDialogs
from .memorization_filter import filter_questions, should_drop, ABSOLUTE_WORDS
from . import settings as settings_mod
from .paths import DATA_DIR
import copy
import os
import random
import re
import sqlite3
import sys
import tkinter as tk
from dataclasses import dataclass, field
from datetime import datetime
from tkinter import filedialog, ttk
from typing import Optional

# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------

APP_TITLE = "刷题"
SAVE_FILE = "quiz_progress.json"          # 与 exe / 脚本同目录
SHEET_NAMES_KEY = "sheet_names"


@dataclass
class Question:
    """一道题"""
    qid: str                       # 唯一 id：sheet__row
    sheet: str                     # 来源 sheet 名
    row: int                       # 在 sheet 中的行号（从 1 计）
    type: str                      # "单选" / "多选" / "判断"
    stem: str                      # 题干
    options: list[str] = field(default_factory=list)   # ["A. xxx", "B. xxx", ...]
    answer: list[str] = field(default_factory=list)    # ["A", "B"]
    explanation: str = ""
    difficulty: str = "中等"

    @property
    def key(self) -> str:
        """用户作答的 key（去空格字母）"""
        return "".join(self.answer)


# ---------------------------------------------------------------------------
# 题库导入
# ---------------------------------------------------------------------------

def _find_option_starts(text: str) -> list:
    """两遍匹配选项起点：
    1) 严格模式（字母前是行首/空白/中文/标点）——正常格式都命中这个；
    2) 宽松模式（字母前允许是数字，兼容 A.1024B.1000 无空格写法）。
    严格模式能切出 >=2 个选项且首选项从头开始则用严格，否则退回宽松。
    """
    strict_re = re.compile(
        r"(?:^|(?<=[\u4e00-\u9fff\s;；,，。！？!？:：、）)】”\"']))"
        r"([A-Za-z])\s*[.、．:：\)）]"
    )
    loose_re = re.compile(
        r"(?:^|(?<=[\u4e00-\u9fff0-9\s;；,，。！？!？:：、）)】”\"']))"
        r"([A-Za-z])\s*[.、．:：\)）]"
    )
    strict = list(strict_re.finditer(text))
    if len(strict) >= 2 and not text[:strict[0].start()].strip():
        return strict
    return list(loose_re.finditer(text))


def parse_options(raw: str) -> list[str]:
    """从选项单元格切分出选项列表，如 'A. 选项一 B. 选项二' -> ['A. 选项一', 'B. 选项二']。
    兼容写法：
      - 空格分隔：  A. xx B. xx
      - 无空格：    A.xxxB.xxx
      - 换行分隔：  每行一个选项
      - 分隔符：    A. / A、 / A: / A) / A）
    解析失败返回 []。
    """
    if raw is None:
        return []
    text = str(raw).strip()
    if not text:
        return []
    # 归一化横向空白（保留换行）
    text = re.sub(r"[ \t\u3000]+", " ", text)
    starts = _find_option_starts(text)
    if not starts:
        return []
    # 第一个选项必须从头开始，否则说明整段不是选项格式
    if text[:starts[0].start()].strip():
        return []
    opts: list[str] = []
    seen: set[str] = set()
    for i, m in enumerate(starts):
        end = starts[i + 1].start() if i + 1 < len(starts) else len(text)
        body = re.sub(r"\s+", " ", text[m.end():end]).strip().rstrip("；;，,、 ")
        if not body:
            continue
        letter = m.group(1).upper()
        if letter in seen:
            continue
        seen.add(letter)
        opts.append(f"{letter}. {body}")
    return opts
# 拆分列式选项：如 选项A / 选项B / 选项1 / 选项一
_SEP_OPTION_RE = re.compile(r"^选项\s*([A-Ha-h]|\d{1,2}|[一二三四五六七八九十]+)$")
_CN_NUM = {"一": "A", "二": "B", "三": "C", "四": "D", "五": "E",
           "六": "F", "七": "G", "八": "H", "九": "I", "十": "J"}


def _letter_for(n: str) -> Optional[str]:
    """选项列序号 -> 字母。'A'->'A'，'1'->'A'，'一'->'A'，不支持返回 None。"""
    n = n.strip()
    if n in _CN_NUM:
        return _CN_NUM[n]
    if n.isdigit():
        k = int(n)
        return chr(ord("A") + k - 1) if 1 <= k <= 26 else None
    return n.upper()


def find_separate_option_cols(header: list[str]) -> list[tuple[str, int]]:
    """找出表头中拆开的选项列（选项A/选项B/选项1/选项一…），按字母顺序返回 [(letter, idx), ...]。"""
    found: list[tuple[str, int]] = []
    for i, h in enumerate(header):
        if not h:
            continue
        m = _SEP_OPTION_RE.match(h.strip())
        if m:
            letter = _letter_for(m.group(1))
            if letter:
                found.append((letter, i))
    found.sort(key=lambda t: t[0])
    return found


def parse_answer(raw: str, qtype: str) -> list[str]:
    """解析答案字符串 -> ['A', 'B'] 形式。
    支持 'A' / 'A,B' / 'A B' / 'AB' / 'a,b' 等写法。
    """
    if raw is None:
        return []
    text = str(raw).strip().upper()
    if not text:
        return []
    # 判断题：答案可能是 对/错、√/×、T/F、1/0
    if qtype == "判断":
        mapping = {"对": "A", "T": "A", "1": "A", "√": "A", "正确": "A",
                   "错": "B", "F": "B", "0": "B", "×": "B", "错误": "B"}
        for k, v in mapping.items():
            if k in text:
                return [v]
        # 否则按字母处理
    letters = re.findall(r"[A-Za-z]", text)
    # 去重且保持顺序
    seen, out = set(), []
    for ch in letters:
        c = ch.upper()
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def _normalize_header(h) -> str:
    """表头归一化：去空白（'题 目' -> '题目'）、去掉括号注释（'题干（必填）' -> '题干'、'选项E(勿删)' -> '选项E'）。"""
    if h is None:
        return ""
    s = re.sub(r"\s+", "", str(h)).strip()
    s = re.sub(r"[（(][^）)]*[）)]", "", s)  # 去掉（必填）(勿删) 等注释
    return s


# 表头别名：标准名 -> 可识别的写法（归一化后比较）
_HEADER_ALIASES = {
    "题目": {"题目", "题干", "问题", "题目描述", "题干内容", "question"},
    "题型": {"题型", "题目类型", "问题类型", "类型", "qtype"},
    "选项": {"选项", "选项内容", "options"},
    "答案": {"答案", "正确答案", "标准答案", "参考答案", "正确选项", "answer"},
    "解析": {"解析", "答案解析", "分析", "解释", "说明", "explanation"},
    "难度": {"难度", "难易", "难度等级", "等级", "difficulty"},
}


def _resolve_header(name: str) -> str:
    """把归一化后的列名映射到标准名（题目/题型/选项/答案/解析/难度），不认识则原样返回。"""
    n = name.strip().lower()
    for std, aliases in _HEADER_ALIASES.items():
        if n in {a.lower() for a in aliases}:
            return std
    return name


def load_excel(path: str) -> tuple[dict[str, list[Question]], list[str]]:
    """从 Excel 文件加载题库。
    返回 ({sheet名: [Question, ...]}, 跳过行提示列表)。
    支持的表头（顺序无关，缺省列用默认值）：
      题型  题目*  选项 / 选项A..选项H(或 选项1..选项10)  答案  解析  难度
    遇到格式问题会抛出带中文说明的 ValueError。
    """
    try:
        import openpyxl
    except ImportError:
        raise ValueError(
            "未安装 openpyxl，请先执行：\n    pip install openpyxl"
        )

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    result: dict[str, list[Question]] = {}
    skipped: list[str] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        # 第 1 行为表头（归一化：去空格、去括号注释，再按别名映射到标准名）
        header = [_normalize_header(c) for c in rows[0]]
        if not any(header):
            continue

        # 建立标准列名 -> 索引映射（题干->题目、正确答案->答案 等别名自动识别）
        col_idx: dict[str, int] = {}
        for i, h in enumerate(header):
            if not h:
                continue
            std = _resolve_header(h)
            if std not in col_idx:
                col_idx[std] = i

        if "题目" not in col_idx:
            missing = "、".join(h for h in header if h)
            skipped.append(f"sheet「{sheet_name}」：未找到题目/题干列，已跳过（实际表头：{missing}）")
            continue

        # 选项列：单独的「选项」列，或拆开的 选项A/选项B/选项1/选项一 多列
        sep_opt_cols = find_separate_option_cols(header)
        has_option_col = "选项" in col_idx

        def cell(row_vals, col_name, default=None):
            idx = col_idx.get(col_name)
            if idx is None or idx >= len(row_vals):
                return default
            v = row_vals[idx]
            return v if v is not None else default

        def build_options(row_vals) -> list[str]:
            """优先用单独「选项」列；否则从拆开的多列拼接。"""
            if has_option_col:
                opts = parse_options(str(cell(row_vals, "选项", "") or ""))
                if opts:
                    return opts
            if sep_opt_cols:
                # 拆分列：每列是一个选项的文本（不带字母前缀）
                opts = []
                for letter, idx in sep_opt_cols:
                    val = row_vals[idx] if idx < len(row_vals) else None
                    if val is None:
                        continue
                    body = re.sub(r"\s+", " ", str(val)).strip()
                    if not body:
                        continue
                    # 若单元格自带 "A. xxx" 前缀且字母与本列一致，去掉冗余前缀
                    with_prefix = parse_options(body)
                    if with_prefix and len(with_prefix) == 1 and with_prefix[0][0] == letter:
                        opts.append(with_prefix[0])
                    else:
                        opts.append(f"{letter}. {body}")
                return opts
            return []

        questions: list[Question] = []
        for rownum, row_vals in enumerate(rows[1:], start=2):
            if row_vals is None:
                continue
            stem = cell(row_vals, "题目", "")
            if not stem or not str(stem).strip():
                continue  # 空行
            qtype = str(cell(row_vals, "题型", "单选") or "单选").strip().replace(" ", "")
            if qtype in ("单选题", "单项选择题"):
                qtype = "单选"
            elif qtype in ("多选题", "多项选择题"):
                qtype = "多选"
            elif qtype in ("判断题", "对错题"):
                qtype = "判断"
            if qtype not in ("单选", "多选", "判断"):
                qtype = "单选"

            options = build_options(row_vals)
            if qtype == "判断":
                # 判断题：若给了真实选项列则用之，否则默认 A.正确 B.错误
                options = options or ["A. 正确", "B. 错误"]

            ans_raw = cell(row_vals, "答案", "")
            answer = parse_answer(ans_raw, qtype)
            # 判断题答案可能直接写中文（对/错），parse_answer 已映射为字母
            # 若答案字母超出选项范围，按「选项文本匹配」兜底
            opt_letters = {o[0] for o in options}
            answer = [a for a in answer if a in opt_letters]
            if not answer and ans_raw and qtype != "判断":
                # 兜底：答案写的是选项文本（如 "主板"），尝试按文本匹配
                ans_text = str(ans_raw).strip()
                for o in options:
                    body = re.sub(r"^[A-Za-z]\s*[.、．:：\)）]\s*", "", o).strip()
                    if body and (body == ans_text or ans_text in body):
                        answer.append(o[0])
                        break

            if not options or not answer:
                reason = "缺少选项" if not options else "缺少有效答案"
                skipped.append(f"sheet「{sheet_name}」第{rownum}行：{reason}（{str(stem)[:20]}…），已跳过")
                continue

            explanation = str(cell(row_vals, "解析", "") or "").strip()
            difficulty = str(cell(row_vals, "难度", "中等") or "中等").strip() or "中等"

            questions.append(Question(
                qid=f"{sheet_name}__{rownum}",
                sheet=sheet_name,
                row=rownum,
                type=qtype,
                stem=str(stem).strip(),
                options=options,
                answer=answer,
                explanation=explanation,
                difficulty=difficulty,
            ))

        if questions:
            result[sheet_name] = questions

    wb.close()
    return result, skipped


# ---------------------------------------------------------------------------
# 长期题库数据库（SQLite）+ 会话（JSON）
# ---------------------------------------------------------------------------

def _base_dir() -> str:
    """数据文件目录：仓库根下的 data/（打包后为 exe 同目录的 data/）"""
    return DATA_DIR


DB_FILE = os.path.join(DATA_DIR, "quiz_bank.db")
SESSION_FILE = os.path.join(DATA_DIR, "quiz_session.json")

# ---------- 现代配色（扁平 + 少量阴影感的分层色） ----------
FONT_FAMILY = "Microsoft YaHei UI"           # Windows 现代无衬线字体

# 字体缩放系数（0.8–1.5），由综合设置写入；_font() 用它做全局缩放
FONT_SCALE = 1.0


def _font(size, weight="normal"):
    """统一的字体工厂：Microsoft YaHei UI（无该字体时 Tk 会回退到默认）。

    size 是逻辑字号，乘上 FONT_SCALE 得到实际字号，实现全局缩放。
    """
    return (FONT_FAMILY, max(1, round(size * FONT_SCALE)), weight)


# 背景 / 表面
BG_APP = "#ffffff"        # 应用整体背景（浅灰）
BG_CARD = "#ffffff"       # 卡片 / 表面（白）
BG_SUBTLE = "#f7f7f8"     # 次级面板（更浅的灰白）
BORDER = "#e5e7eb"        # 描边 / 分隔线
BORDER_SOFT = "#eef0f3"

# 文字
FG_TEXT = "#202123"       # 主文字（深灰黑）
FG_MUTED = "#707078"      # 次要文字
FG_FAINT = "#9ca3af"      # 弱化文字
FG_ON_ACCENT = "#ffffff"  # 主题色上的文字

# 主题色（深空黑，可由综合设置切换；hover/active/soft/on 由 _apply_theme 派生）
ACCENT = "#242424"
ACCENT_HOVER = "#3a3a3a"
ACCENT_ACTIVE = "#111111"
ACCENT_SOFT = "#ececed"   # 主题色浅底（选中/高亮背景）

# 语义色
C_OK = "#16a34a"          # 答对（绿）
C_OK_SOFT = "#dcfce7"     # 答对浅底
C_BAD = "#dc2626"         # 答错（红）
C_BAD_SOFT = "#fee2e2"    # 答错浅底
C_UNANSWERED = "#eeeeef"  # 未答（灰）
C_CURRENT = ACCENT        # 当前题（主题色）
C_TEXT_ON_CELL = "#ffffff"  # 深色格子上的序号
C_TEXT_ON_LIGHT = "#374151"  # 浅色格子上的序号


class BankDB:
    """长期题库：导入的试卷/题目永久保存在 SQLite，错题本也存这里。"""

    def __init__(self, path: str = DB_FILE):
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.execute(
            """CREATE TABLE IF NOT EXISTS questions(
                 id INTEGER PRIMARY KEY AUTOINCREMENT,
                 sheet TEXT NOT NULL,
                 row INTEGER NOT NULL,
                 type TEXT, stem TEXT, options TEXT, answer TEXT,
                 explanation TEXT, difficulty TEXT, source TEXT,
                 seq INTEGER DEFAULT 0,
                 UNIQUE(sheet, row))"""
        )
        self.conn.execute(
            """CREATE TABLE IF NOT EXISTS wrong(
                 qid TEXT PRIMARY KEY,        -- sheet__row
                 first_seen TEXT, last_wrong TEXT)"""
        )
        self.conn.execute("CREATE TABLE IF NOT EXISTS favorites(qid TEXT PRIMARY KEY, saved_at TEXT NOT NULL)")
        self.conn.commit()
        # 兼容旧库：补 seq 列
        cols = [r[1] for r in self.conn.execute("PRAGMA table_info(questions)")]
        if "seq" not in cols:
            self.conn.execute("ALTER TABLE questions ADD COLUMN seq INTEGER DEFAULT 0")
            self.conn.commit()

    # ---- 题库 ----
    def upsert_sheet(self, sheet: str, source: str,
                     questions: list[Question], seq: int = 0) -> None:
        """写入/覆盖一个 sheet 的整份题库"""
        with self.conn:
            self.conn.execute("DELETE FROM favorites WHERE qid IN (SELECT sheet || '__' || row FROM questions WHERE sheet=?)", (sheet,))
            self.conn.execute("DELETE FROM questions WHERE sheet = ?", (sheet,))
            self.conn.executemany(
                """INSERT OR REPLACE INTO questions
                   (sheet, row, type, stem, options, answer, explanation, difficulty, source, seq)
                   VALUES (?,?,?,?,?,?,?,?,?,?)""",
                [(q.sheet, q.row, q.type, q.stem,
                  json.dumps(q.options, ensure_ascii=False),
                  ",".join(q.answer), q.explanation, q.difficulty, source, seq)
                 for q in questions],
            )
        # 错题本里已不在题库中的 qid 清掉
        with self.conn:
            self.conn.execute(
                "DELETE FROM wrong WHERE qid NOT IN "
                "(SELECT sheet || '__' || CAST(row AS TEXT) FROM questions)")
        self.conn.commit()

    def get_sheet(self, sheet: str) -> list[Question]:
        rows = self.conn.execute(
            "SELECT sheet,row,type,stem,options,answer,explanation,difficulty "
            "FROM questions WHERE sheet = ? ORDER BY row", (sheet,)).fetchall()
        out = []
        for r in rows:
            out.append(Question(
                qid=f"{r[0]}__{r[1]}", sheet=r[0], row=r[1], type=r[2] or "单选",
                stem=r[3] or "",
                options=json.loads(r[4] or "[]"),
                answer=(r[5] or "").split(",") if r[5] else [],
                explanation=r[6] or "", difficulty=r[7] or "中等",
            ))
        return out

    def get_question(self, qid: str) -> Optional[Question]:
        m = re.match(r"^(.+)__(\d+)$", qid)
        if not m:
            return None
        r = self.conn.execute(
            "SELECT sheet,row,type,stem,options,answer,explanation,difficulty "
            "FROM questions WHERE sheet=? AND row=?", (m.group(1), int(m.group(2)))).fetchone()
        if not r:
            return None
        return Question(qid=qid, sheet=r[0], row=r[1], type=r[2] or "单选",
                        stem=r[3] or "", options=json.loads(r[4] or "[]"),
                        answer=(r[5] or "").split(",") if r[5] else [],
                        explanation=r[6] or "", difficulty=r[7] or "中等")

    def sheet_names(self) -> list[str]:
        # 按导入顺序（seq）返回 sheet
        return [r[0] for r in self.conn.execute(
            "SELECT sheet FROM questions GROUP BY sheet ORDER BY MIN(seq)")]

    def total_count(self) -> int:
        return self.conn.execute("SELECT COUNT(*) FROM questions").fetchone()[0]

    def remove_sheet(self, sheet: str) -> None:
        with self.conn:
            self.conn.execute("DELETE FROM favorites WHERE qid IN (SELECT sheet || '__' || row FROM questions WHERE sheet=?)", (sheet,))
            self.conn.execute("DELETE FROM questions WHERE sheet = ?", (sheet,))
            # 错题本中属于该 sheet 的 qid（形如 sheet__row）一并清除
            self.conn.execute(
                "DELETE FROM wrong WHERE qid LIKE ?", (sheet + "__%",))

    def remove_all(self) -> None:
        """清空整个题库（所有 sheet + 错题本）"""
        with self.conn:
            self.conn.execute("DELETE FROM favorites")
            self.conn.execute("DELETE FROM questions")
            self.conn.execute("DELETE FROM wrong")

    # ---- 错题本（长期） ----
    def wrong_add(self, qid: str) -> None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        with self.conn:
            self.conn.execute(
                "INSERT INTO wrong(qid, first_seen, last_wrong) VALUES(?,?,?) "
                "ON CONFLICT(qid) DO UPDATE SET last_wrong=excluded.last_wrong",
                (qid, now, now))

    def wrong_remove(self, qid: str) -> None:
        with self.conn:
            self.conn.execute("DELETE FROM wrong WHERE qid=?", (qid,))

    def wrong_list(self, sheet: Optional[str] = None) -> list[Question]:
        sql = ("SELECT w.qid FROM wrong w JOIN questions q "
               "ON w.qid = q.sheet || '__' || q.row")
        params = ()
        if sheet is not None:
            sql += " WHERE q.sheet=?"
            params = (sheet,)
        sql += " ORDER BY w.last_wrong DESC"
        qids = [r[0] for r in self.conn.execute(sql, params)]
        out = []
        for qid in qids:
            q = self.get_question(qid)
            if q:
                out.append(q)
        return out

    def is_favorite(self, qid):
        return self.conn.execute("SELECT 1 FROM favorites WHERE qid=?", (qid,)).fetchone() is not None

    def set_favorite(self, qid, saved):
        with self.conn:
            if saved:
                if self.get_question(qid):
                    self.conn.execute("INSERT OR IGNORE INTO favorites VALUES (?, ?)", (qid, datetime.now().isoformat()))
            else:
                self.conn.execute("DELETE FROM favorites WHERE qid=?", (qid,))

    def favorite_list(self, sheet: Optional[str] = None):
        # 删除题库后清理无对应题目的收藏，避免下次导入时恢复失效收藏。
        with self.conn:
            self.conn.execute("DELETE FROM favorites WHERE qid NOT IN (SELECT sheet || '__' || row FROM questions)")
        sql = ("SELECT f.qid FROM favorites f JOIN questions q "
               "ON f.qid = q.sheet || '__' || q.row")
        params = ()
        if sheet is not None:
            sql += " WHERE q.sheet=?"
            params = (sheet,)
        sql += " ORDER BY f.saved_at DESC, f.qid"
        return [self.get_question(row[0]) for row in self.conn.execute(sql, params)]

    def close(self) -> None:
        self.conn.close()


# ---- 会话（上次刷题位置，重启可续） ----

def load_session(path: str = SESSION_FILE) -> dict:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def save_session(path: str, data: dict) -> None:
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except OSError:
        pass


def clear_session(path: str = SESSION_FILE) -> None:
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


def init_progress(data: dict, excel_path: str, sheet_names: list[str]) -> dict:
    """(兼容旧调用，现无实际作用)"""
    return data


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------

class QuizApp(AIExplanationMixin, ShortcutMixin):
    def __init__(self, root: tk.Tk):
        self.root = root
        self.dialogs = AppDialogs(root, FONT_FAMILY)
        root.title(APP_TITLE)
        root.geometry("1400x820")
        root.minsize(1100, 600)

        # 长期数据库
        self.db = BankDB()

        # 会话状态
        self.sheets: dict[str, list[Question]] = {}   # 当前从 DB 加载的题库
        self.sheet_names: list[str] = []
        self.current_sheet: Optional[str] = None
        self.mode: str = "normal"                     # "normal" | "wrong"
        self.current_qs: list[Question] = []
        self.order: list[str] = []                    # 本轮作答顺序（qid 列表）
        self.pos: int = 0
        self.answers: dict[str, dict] = {}            # qid -> {"sel": [...], "ok": bool}
        self.score_correct: int = 0
        self.score_wrong: int = 0
        self._restored = False
        self._pending: Optional[int] = None      # 待触发的自动跳转定时器
        self._summary_mode = False
        self._normal_session = None
        self._sheet_sessions = {}
        self._wrong_sessions = {}
        self.qmap: dict[str, Question] = {}      # qid -> Question

        self._init_shortcuts(os.path.join(_base_dir(), "shortcut_settings.json"))
        self._init_settings(os.path.join(_base_dir(), "app_settings.json"))
        self.dialogs.set_palette_provider(lambda: {
            "BG_CARD": BG_CARD, "FG_TEXT": FG_TEXT,
            "FG_MUTED": FG_MUTED, "ACCENT": ACCENT,
        })
        self._build_ui()
        self._install_shortcut_bindings()
        self._init_ai(os.path.join(_base_dir(), "ai_settings.json"))
        self._load_from_db()
        # 加载上次的会话进度（不自动进入做题页，只在导航页显示各卷进度）
        self._load_sessions_only()
        # 重新渲染导航页，让各卷卡片反映刚加载的题库与进度
        self._render_bank()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._theme_watch_id = self.root.after(2000, self._watch_system_theme)

    # ---------------- UI 构建 ----------------

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam" if "clam" in style.theme_names() else "default")

        # 窗口整体浅灰背景，白卡片浮于其上
        self.root.config(bg=BG_APP)
        self.root.option_add("*Font", _font(10))
        style.configure("TFrame", background=BG_CARD)
        style.configure("TLabel", background=BG_CARD, foreground=FG_TEXT, font=_font(10))
        style.configure("Treeview", font=_font(11), rowheight=40, background=BG_CARD,
                        fieldbackground=BG_CARD, foreground=FG_TEXT, borderwidth=0)
        style.configure("Treeview.Heading", font=_font(10, "bold"), background=BG_SUBTLE)
        style.map("Treeview", background=[("selected", ACCENT_SOFT)], foreground=[("selected", FG_TEXT)])

        # --- 通用按钮（扁平、圆角感靠 padding + 无边框） ---
        style.configure("TButton", font=_font(10), padding=(12, 7), borderwidth=0,
                        relief="flat", background=BG_CARD, foreground=FG_TEXT,
                        lightcolor=BG_CARD, darkcolor=BG_CARD,
                        focuscolor=ACCENT_SOFT, focusthickness=0)
        style.map("TButton",
                  background=[("active", BG_SUBTLE), ("disabled", BG_SUBTLE)],
                  foreground=[("disabled", FG_FAINT)])
        # 主题按钮（靛蓝）
        style.configure("Accent.TButton", font=_font(10, "bold"), padding=(14, 8),
                        borderwidth=0, relief="flat", background=ACCENT,
                        foreground=FG_ON_ACCENT, lightcolor=ACCENT, darkcolor=ACCENT,
                        focuscolor=ACCENT_SOFT, focusthickness=0)
        style.map("Accent.TButton",
                  background=[("active", ACCENT_HOVER), ("pressed", ACCENT_ACTIVE),
                             ("disabled", BG_SUBTLE)],
                  foreground=[("disabled", FG_FAINT)])
        # 次级/幽灵按钮（描边感）
        style.configure("Soft.TButton", font=_font(10), padding=(12, 7),
                        borderwidth=1, relief="flat", background=BG_CARD,
                        foreground=FG_TEXT, lightcolor=BORDER, darkcolor=BORDER,
                        focuscolor=ACCENT_SOFT, focusthickness=0)
        style.map("Soft.TButton", background=[("active", ACCENT_SOFT), ("disabled", BG_SUBTLE)],
                  foreground=[("disabled", FG_FAINT)])
        # 删除按钮（灰色弱化，避免过于显眼）
        style.configure("Danger.TButton", font=_font(9), padding=(10, 5),
                        borderwidth=1, relief="flat", background=BG_CARD,
                        foreground=FG_MUTED, lightcolor=BORDER, darkcolor=BORDER,
                        focuscolor=BG_SUBTLE, focusthickness=0)
        style.map("Danger.TButton",
                  background=[("active", BG_SUBTLE), ("disabled", BG_SUBTLE)],
                  foreground=[("active", FG_MUTED), ("disabled", FG_FAINT)])

        # --- 标签 ---
        style.configure("Header.TLabel", font=_font(18, "bold"), background=BG_APP,
                        foreground=FG_TEXT)
        style.configure("Sub.TLabel", font=_font(10), background=BG_APP, foreground=FG_MUTED)
        style.configure("SubCard.TLabel", font=_font(10), background=BG_SUBTLE,
                        foreground=FG_MUTED)
        style.configure("Question.TLabel", font=_font(16), background=BG_CARD,
                        foreground=FG_TEXT, wraplength=620, justify="left")
        style.configure("Explain.TLabel", font=_font(11), background=BG_SUBTLE,
                        foreground=FG_MUTED, wraplength=620, justify="left",
                        padding=(12, 10))
        style.configure("Ok.TLabel", font=_font(11, "bold"), background=C_OK_SOFT,
                        foreground=C_OK, padding=(12, 10))
        style.configure("Bad.TLabel", font=_font(11, "bold"), background=C_BAD_SOFT,
                        foreground=C_BAD, padding=(12, 10))
        style.configure("Option.TCheckbutton", font=_font(11), background=BG_CARD,
                        foreground=FG_TEXT, padding=(8, 4))

        # --- 下拉框 ---
        style.configure("TCombobox", fieldbackground=BG_CARD, background=BG_CARD,
                        foreground=FG_TEXT, arrowcolor=ACCENT, bordercolor=BORDER,
                        lightcolor=BG_CARD, darkcolor=BG_CARD, selectbackground=ACCENT,
                        selectforeground=FG_ON_ACCENT, padding=5)
        style.map("TCombobox", fieldbackground=[("readonly", BG_CARD), ("disabled", BG_SUBTLE)],
                  arrowcolor=[("readonly", ACCENT)], bordercolor=[("focus", ACCENT)])

        # --- 进度条（细、主题色） ---
        style.configure("TProgressbar", background=ACCENT, troughcolor=BG_SUBTLE,
                        bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT,
                        thickness=6)
        style.map("TProgressbar", background=[("disabled", BORDER)])

        # --- 滚动条（细、浅灰） ---
        style.configure("Vertical.TScrollbar", background=BORDER, troughcolor=BG_APP,
                        bordercolor=BG_APP, arrowcolor=FG_MUTED, relief="flat",
                        arrowsize=12)
        style.map("Vertical.TScrollbar",
                  background=[("active", FG_FAINT), ("pressed", FG_FAINT)])

        # 顶栏：白底头部卡片（标题 + 题库信息 + 控件）
        top = tk.Frame(self.root, bg=BG_CARD)
        self.quiz_header = top
        # 左侧标题区
        head = tk.Frame(top, bg=BG_CARD)
        head.pack(side="left", padx=(8, 32), pady=8)
        tk.Label(head, text="刷题", bg=BG_CARD, fg=FG_TEXT,
                 font=_font(14, "bold")).pack(anchor="w")
        #self.file_label = tk.Label(head, text="尚未导入题库", bg=BG_CARD,
        #                           fg=FG_MUTED, font=_font(10), anchor="w",
        #                           justify="left")
        #self.file_label.pack(anchor="w", pady=(3, 0))

        # 顺序选择：随机乱序 / 原顺序
        ctl = tk.Frame(top, bg=BG_CARD)
        ctl.pack(side="left", fill="x", expand=True, pady=8)
        tk.Label(ctl, text="顺序", bg=BG_CARD, fg=FG_MUTED, font=_font(10)).pack(side="left", padx=(6, 6))
        self.order_var = tk.StringVar(value="随机乱序")
        self.order_combo = ttk.Combobox(ctl, textvariable=self.order_var,
                                        state="disabled", width=9,
                                        values=["随机乱序", "原顺序", "题型分组"])
        self.order_combo.pack(side="left", padx=(0, 14))
        self.order_combo.bind("<<ComboboxSelected>>", self._on_order_change)

        # 试卷选择
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(ctl, textvariable=self.sheet_var,
                                        state="disabled", width=22)
        self.sheet_combo.pack(side="left", padx=(0, 10))
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_change)
        ttk.Button(ctl, text="导入题库", command=self.on_import,
                   style="Accent.TButton").pack(side="left")

        # ============ 两个顶层视图：导航页（我的题库）/ 做题页 ============
        # 做题页：进度条 + 题目卡(答题) + 底部操作栏，整体装进 quiz_view
        self.quiz_view = tk.Frame(self.root, bg=BG_APP)
        # 导航页：我的题库（卡片列表），装进 bank_view
        self.bank_view = tk.Frame(self.root, bg=BG_APP)
        bank_sidebar = tk.Frame(self.bank_view, bg=BG_SUBTLE, width=200, padx=16, pady=28)
        bank_sidebar.pack(side="left", fill="y")
        bank_sidebar.pack_propagate(False)
        tk.Label(bank_sidebar, text="刷题空间", bg=BG_SUBTLE, fg=FG_TEXT,
                 font=_font(14, "bold")).pack(anchor="w", pady=(0, 32))
        for label, action in (("＋  导入题库", self.on_import),
                              ("我的题库", self.show_bank)):
            ttk.Button(bank_sidebar, text=label, command=action, style="Soft.TButton").pack(fill="x", pady=5)
        ttk.Button(bank_sidebar, text="设置", command=self.open_settings,
                   style="Soft.TButton").pack(side="bottom", fill="x")

        # --- 做题页内部：进度条（细、主题色，独立一行） ---
        self.progress_var = tk.IntVar(value=0)
        prog = tk.Frame(self.quiz_view, bg=BG_APP)
        prog.pack(fill="x", padx=28, pady=(0, 12))
        self.progress_bar = ttk.Progressbar(prog, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", side="left", expand=True)
        self.progress_text = tk.Label(prog, text="0 / 0", bg=BG_APP, fg=FG_MUTED,
                                      font=_font(10), anchor="e")
        self.progress_text.pack(side="left", padx=(10, 0))

        # 中部：题目卡（左：答题卡 / 中：题目 / 右：AI 解析），白底卡片
        self.card = tk.Frame(self.quiz_view, bg=BG_CARD, highlightthickness=0,
                             highlightbackground=BORDER)
        self.card.pack(fill="both", expand=True, padx=24, pady=(0, 16))
        self.card.columnconfigure(0, weight=4)
        self.card.columnconfigure(1, weight=1)
        self.card.rowconfigure(0, weight=1)
        # 卡片内边距：用一个内层 frame 撑起 padding
        inner = tk.Frame(self.card, bg=BG_CARD, padx=0, pady=8)
        inner.grid(row=0, column=0, columnspan=2, sticky="nsew")
        inner.columnconfigure(0, weight=0)
        inner.columnconfigure(1, weight=3, minsize=330)
        inner.columnconfigure(2, weight=2, minsize=300)
        inner.rowconfigure(0, weight=1)

        # --- 右栏：题目内容（装进滚动区，长题干不再撑大窗口） ---
        left = tk.Frame(inner, bg=BG_CARD)
        left.grid(row=0, column=1, sticky="nsew", padx=(20, 0))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)   # 行0=题目滚动区（可变高），行1=固定底部条
        left.rowconfigure(1, weight=0)

        # 题目区：白卡 + 滚动（题干/选项/提交/结果），过长时整段下移滚动
        self.q_canvas = tk.Canvas(left, highlightthickness=0, bg=BG_CARD)
        self.q_canvas.grid(row=0, column=0, sticky="nsew")
        self.q_sb = ttk.Scrollbar(left, orient="vertical", command=self.q_canvas.yview)
        self.q_sb.grid(row=0, column=1, sticky="ns")
        self.q_canvas.config(yscrollcommand=self.q_sb.set)

        self.q_content = tk.Frame(self.q_canvas, bg=BG_CARD, padx=12, pady=8)
        self.q_win = self.q_canvas.create_window((0, 0), window=self.q_content, anchor="nw")
        self.q_content.columnconfigure(0, weight=1)
        self.q_content.bind(
            "<Configure>",
            lambda e: self.q_canvas.config(scrollregion=self.q_canvas.bbox("all")))
        self.q_canvas.bind("<Configure>", self._fit_q_canvas)

        # 固定底部条：解析 + 上一题/下一题 永远贴底，不随题目长短跳动
        self.q_bottom = tk.Frame(self.q_content, bg=BG_CARD, padx=0, pady=10)
        self.q_bottom.grid(row=5, column=0, sticky="ew", padx=4)
        self.q_bottom.columnconfigure(0, weight=1)

        self.placeholder = tk.Label(
            self.q_content,
            text="点击「导入题库」选择 Excel 文件开始刷题\n\n"
                 "支持题型：单选 / 多选 / 判断\n"
                 "单选/判断点选项即判分，多选点「提交本题」\n"
                 "答对后自动下一题，答错停留本题看解析\n"
                 "可选随机乱序 / 原顺序 / 题型分组（单选→多选→判断）\\n"
                 "导入一次后题库会永久保存，下次打开自动恢复",
            bg=BG_CARD, fg=FG_MUTED, font=_font(11), justify="left",
            anchor="center",
        )
        self.placeholder.grid(row=0, column=0, columnspan=2, sticky="nsew")

        self.question_meta = tk.Label(self.q_content, bg=BG_CARD, fg=ACCENT, font=_font(10))
        self.question_label = ttk.Label(self.q_content, style="Question.TLabel",
                                        font=_font(16), justify="left", anchor="nw",
                                        text="")
        self.options_frame = ttk.Frame(self.q_content)
        try:  # 去掉 clam 主题 ttk.Frame 自带约 2px 边距，让选项行贴齐卡片内容边缘
            self.options_frame.configure(padding=0, borderwidth=0, relief="flat")
        except tk.TclError:
            pass
        # 多选题提交按钮：放在选项正下方（题目区内，不再放底部操作栏）
        self.btn_submit = ttk.Button(self.q_content, text="提交答案",
                                     style="Accent.TButton", command=self.on_submit)
        # 结果条：对错 + 正确答案（内联，不弹窗）
        self.result_label = ttk.Label(self.q_content, text="", style="Ok.TLabel")
        self.explain_label = ttk.Label(self.q_bottom, style="Explain.TLabel", text="")
        # 内容控件高度变化（题干换行、选项增减、解析出现）→ 重同步滚动区
        for w in (self.question_label, self.options_frame, self.btn_submit,
                  self.result_label, self.explain_label):
            w.bind("<Configure>", lambda e: self.root.after_idle(self._sync_q_scroll))
        # 题目四件套延迟到开卷时才 grid（避免与占位提示占同一格被挤掉）
        self._q_shown = False
        self._last_q_w = 0

        # --- 左栏：题目矩阵（序号+状态色，点击跳转），浅灰面板 ---
        right = tk.Frame(inner, bg=BG_SUBTLE, padx=8, pady=12, highlightthickness=0,
                         highlightbackground=BORDER)
        right.grid(row=0, column=0, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        tk.Label(right, text="答题卡", bg=BG_SUBTLE, fg=FG_MUTED,
                 font=_font(10, "bold")).grid(row=0, column=0, columnspan=2,
                                              sticky="ew", padx=10, pady=(10, 6))

        self.list_canvas = tk.Canvas(right, highlightthickness=0, width=240,
                                     bg=BG_SUBTLE)
        self.list_canvas.grid(row=1, column=0, sticky="nsew")
        self.list_sb = ttk.Scrollbar(right, orient="vertical",
                                     command=self.list_canvas.yview)
        self.list_sb.grid(row=1, column=1, sticky="ns")
        self.list_canvas.config(yscrollcommand=self.list_sb.set)
        # 单 Canvas 矩阵：每题一个矩形+序号文本（避免上千个 Button），点击按坐标命中
        self.list_cells: list = []        # [(rect_id, text_id), ...]
        self._cell_colors: list = []      # 每格当前填充色（变化才重绘）
        self._list_cols = 0
        self._list_rows = 0
        self._cell_w = 42
        self._cell_h = 38
        self.list_canvas.bind("<Configure>", self._fit_list_matrix)
        self.list_canvas.bind("<Button-1>", self._on_matrix_click)
        self.list_canvas.configure(yscrollincrement=self._cell_h)
        self._matrix_wheel_delta = 0
        self.list_canvas.bind("<MouseWheel>", self._on_matrix_wheel)
        self.list_canvas.bind("<Button-4>", self._on_matrix_wheel)
        self.list_canvas.bind("<Button-5>", self._on_matrix_wheel)

        legend = tk.Frame(right, bg=BG_SUBTLE)
        legend.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 10), padx=10)
        for color, text in ((C_BAD, "答错"), (C_OK, "答对"),
                            (C_UNANSWERED, "未答"), (C_CURRENT, "当前")):
            tk.Label(legend, bg=color, width=1, height=1).pack(side="left",
                                                                padx=(0, 3), pady=2)
            tk.Label(legend, text=text, bg=BG_SUBTLE, fg=FG_MUTED,
                     font=_font(9)).pack(side="left", padx=(0, 10))

        # 底部：操作栏（白色），属于做题页
        bottom = tk.Frame(self.quiz_view, bg=BG_APP)
        bottom.pack(fill="x", padx=24, pady=(0, 22))
        # 翻题按钮：放进题目滚动内容区（q_content），与选项同一列、同一左右边距，
        # 保证「上一题/下一题」左边缘与选项行左边缘严格对齐。
        style.configure("Nav.Soft.TButton", font=_font(10), padding=(16, 10))
        style.configure("Nav.Accent.TButton", font=_font(10, "bold"), padding=(16, 10))
        self.nav_frame = tk.Frame(self.q_bottom, bg=BG_CARD)
        self.btn_prev = ttk.Button(self.nav_frame, text="上一题", style="Nav.Soft.TButton",
                                   width=10, command=self.on_prev)
        self.btn_prev.pack(side="left", padx=(0, 12))
        self.btn_next = ttk.Button(self.nav_frame, text="下一题", style="Nav.Accent.TButton",
                                   width=10, command=self.on_next)
        self.btn_next.pack(side="left")
        self.btn_favorite = ttk.Button(self.q_content, text="收藏题目", command=self.on_favorite)
        self.btn_favorite.grid(row=6, column=0, sticky="w", padx=4, pady=10)
        ai_panel = tk.Frame(inner, bg=BG_CARD, padx=16, pady=16,
                            highlightthickness=1, highlightbackground=BORDER)
        ai_heading = tk.Frame(ai_panel, bg=BG_CARD)
        ai_heading.pack(fill="x", pady=(0, 12))
        tk.Label(ai_heading, text="AI 解析", bg=BG_CARD, fg=ACCENT,
                 font=_font(12, "bold")).pack(side="left")
        tk.Label(ai_heading, text="学习参考", bg=BG_CARD, fg=FG_MUTED,
                 font=_font(9)).pack(side="right")
        ai_panel.grid(row=0, column=2, sticky="nsew", padx=(20, 0))
        ai_actions = tk.Frame(ai_panel, bg=BG_CARD)
        ai_actions.pack(fill="x", pady=(0, 12))
        self.btn_ai = ttk.Button(ai_actions, text="AI 解析", style="Accent.TButton", command=self.on_ai_explain)
        self.btn_ai.pack(side="left")
        ttk.Button(ai_actions, text="展开阅读", command=self.expand_ai).pack(side="right")
        self.ai_text = tk.Text(ai_panel, height=18, width=1, wrap="word", font=_font(11),
                               relief="flat", bg=BG_CARD, fg=FG_TEXT, padx=4, pady=6, state="disabled")
        self.ai_text.pack(side="left", fill="both", expand=True)
        ai_scroll = ttk.Scrollbar(ai_panel, command=self.ai_text.yview)
        ai_scroll.pack(side="right", fill="y")
        self.ai_text.config(yscrollcommand=ai_scroll.set)
        ttk.Button(bottom, text="设置", command=self.open_settings).pack(side="right", padx=10)
        self.btn_wrong = ttk.Button(bottom, text="重做错题", style="Soft.TButton",
                                    command=self.on_wrong)
        self.btn_wrong.pack(side="left", padx=10)
        self.btn_restart = ttk.Button(bottom, text="重新刷题", style="Soft.TButton",
                                      command=self.on_restart)
        self.btn_restart.pack(side="left", padx=10)
        self.btn_restart.config(state="disabled")
        self.btn_return = ttk.Button(bottom, text="返回正常练习", style="Accent.TButton",
                                     command=self.on_return)
        self.btn_home = ttk.Button(bottom, text="返回题库", style="Soft.TButton",
                                   command=self.show_bank)
        self.btn_home.pack(side="left", padx=10)
        self.status_label = tk.Label(bottom, text="", bg=BG_APP, fg=FG_MUTED,
                                     font=_font(10))
        self.status_label.pack(side="right")

        # ============ 导航页「我的题库」 ============
        # 顶部标题行：左标题 + 右「清空全部」
        bank_head = tk.Frame(self.bank_view, bg=BG_APP)
        bank_head.pack(fill="x", padx=40, pady=(32, 20))
        bank_title_row = tk.Frame(bank_head, bg=BG_APP)
        bank_title_row.pack(fill="x")
        tk.Label(bank_title_row, text="我的题库", bg=BG_APP, fg=FG_TEXT,
                 font=_font(18, "bold")).pack(side="left")
        #self.btn_clear_all = ttk.Button(bank_title_row, text="清空全部",
        #                                style="Danger.TButton",
        #                                command=self._delete_all)
        #self.btn_clear_all.pack(side="right", anchor="s")
        self.bank_subtitle = tk.Label(bank_head, text="", bg=BG_APP, fg=FG_MUTED,
                                      font=_font(10), anchor="w", justify="left")
        self.bank_subtitle.pack(anchor="w", pady=(4, 0))

        # 可滚动区域（题库多时纵向滚动）
        bank_canvas = tk.Canvas(self.bank_view, highlightthickness=0, bg=BG_APP)
        bank_canvas.pack(fill="both", expand=True, padx=40, pady=(8, 24))
        bank_sb = ttk.Scrollbar(self.bank_view, orient="vertical",
                                command=bank_canvas.yview)
        bank_sb.pack(side="right", fill="y", pady=(8, 16))
        bank_canvas.config(yscrollcommand=bank_sb.set)
        self.bank_body = tk.Frame(bank_canvas, bg=BG_APP)
        self._bank_win = bank_canvas.create_window((0, 0), window=self.bank_body,
                                                   anchor="nw")
        self.bank_body.bind(
            "<Configure>",
            lambda e: bank_canvas.config(scrollregion=bank_canvas.bbox("all")))
        # 关键：把内嵌窗口宽度锁定为 canvas 宽度，否则卡片按内容自然宽度展开
        # （曾出现 1352px > 视口，导致右侧按钮被推出屏幕外）。
        def _sync_body_width(e=None):
            w = bank_canvas.winfo_width()
            if w > 1:
                bank_canvas.itemconfig(self._bank_win, width=w)
        bank_canvas.bind("<Configure>", _sync_body_width)
        # 滚轮：绑到 canvas 上；卡片在 bank_body（canvas 内嵌窗口）里，
        # 鼠标停在卡片上时 wheel 事件被内层 widget 吃掉，
        # 所以渲染后再用 _bind_bank_wheel 把 wheel 绑到 bank_body 及其所有后代
        self.bank_canvas = bank_canvas
        bank_canvas.bind("<MouseWheel>",
                         lambda e: bank_canvas.yview_scroll(-1 * int(e.delta / 120), "units"))
        bank_canvas.bind("<Button-4>", lambda e: bank_canvas.yview_scroll(-1, "units"))
        bank_canvas.bind("<Button-5>", lambda e: bank_canvas.yview_scroll(1, "units"))
        # 点击空白处让内层卡片可点击
        bank_canvas.bind("<Button-1>", lambda e: self._bank_click_through(e))
        self.bank_body.bind("<Button-1>", lambda e: self._bank_click_through(e))

        self.bank_cards = []   # 每张卡片的 frame（重建时清空）

        # 初始显示导航页
        self.bank_view.pack(fill="both", expand=True)
        self._current_view = "bank"
        self._render_bank()

    # ---------------- 视图切换 ----------------

    def _set_view(self, view: str):
        """在「导航页」与「做题页」之间切换"""
        if view == "bank":
            self.quiz_header.pack_forget()
            self.quiz_view.pack_forget()
            self.bank_view.pack(fill="both", expand=True)
            self._render_bank()
        else:
            self.bank_view.pack_forget()
            self.quiz_view.pack(fill="both", expand=True)
            self.quiz_header.pack(fill="x", padx=24, pady=(12, 8), before=self.quiz_view)
        self._current_view = view
        if view == "quiz":
            self.q_canvas.focus_set()

    def _bank_click_through(self, e):
        """让 canvas 内层卡片上的点击正常传递给卡片按钮（不触发滚动）"""
        return

    def _bind_bank_wheel(self, widget):
        """把滚轮/中键滚动手势绑到 widget 及其所有后代，统一滚动 bank_canvas。
        卡片是 canvas 内嵌窗口(bank_body)的子控件，wheel 事件会落在它们身上
        而不是 canvas 上，所以必须把 wheel 绑到内层每个控件。"""
        canvas = self.bank_canvas
        def _scroll(e):
            # 鼠标在卡片上时向上/向下滚动整个题库列表
            canvas.yview_scroll(-1 * int(e.delta / 120), "units")
            return "break"   # 防止内层控件再处理
        widget.bind("<MouseWheel>", _scroll)
        widget.bind("<Button-4>", lambda e: (canvas.yview_scroll(-3, "units"), "break")[1])
        widget.bind("<Button-5>", lambda e: (canvas.yview_scroll(3, "units"), "break")[1])
        for child in widget.winfo_children():
            self._bind_bank_wheel(child)

    def on_favorite(self):
        if not self.order:
            return
        q = self._q_at(self.pos)
        saved = not self.db.is_favorite(q.qid)
        self.db.set_favorite(q.qid, saved)
        self.btn_favorite.config(text="★ 已收藏 · 取消" if saved else "☆ 收藏题目")

    def show_favorites(self, sheet: str):
        self._cancel_auto_next()
        window = tk.Toplevel(self.root)
        window.title(f"{sheet} · 收藏题")
        window.geometry("1000x650")
        window.configure(bg=BG_APP)
        body = tk.Frame(window, bg=BG_APP, padx=20, pady=20)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(1, weight=1)
        title = tk.Label(body, text=f"{sheet} · 收藏题", bg=BG_APP, fg=FG_TEXT, font=_font(18, "bold"))
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        listing = ttk.Treeview(body, columns=("stem",), show="headings", selectmode="browse")
        listing.heading("stem", text="题库 · 题目")
        listing.column("stem", width=300)
        listing.grid(row=1, column=0, sticky="ns")
        list_scroll = ttk.Scrollbar(body, command=listing.yview)
        list_scroll.grid(row=1, column=0, sticky="nse")
        listing.config(yscrollcommand=list_scroll.set)
        detail_frame = ttk.Frame(body)
        detail_frame.grid(row=1, column=1, sticky="nsew", padx=(16, 0))
        detail = tk.Text(detail_frame, wrap="word", width=1, font=_font(12), relief="flat", padx=16, pady=16)
        detail.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(detail_frame, command=detail.yview)
        scroll.pack(side="right", fill="y")
        detail.config(yscrollcommand=scroll.set)
        def select(event=None):
            selection = listing.selection()
            q = self.db.get_question(selection[0]) if selection else None
            value = "暂无收藏。做题时点击「收藏题目」，即可在这里查看。"
            if q:
                value = f"{q.sheet} · {q.type} · {q.difficulty}\n\n{q.stem}\n\n" + "\n\n".join(q.options) + "\n\n参考答案：" + "、".join(q.answer) + "\n\n解析：" + (q.explanation or "暂无题库解析")
            detail.config(state="normal")
            detail.delete("1.0", "end")
            detail.insert("1.0", value)
            detail.config(state="disabled")
            remove.config(state="normal" if q else "disabled")
        def refresh():
            listing.delete(*listing.get_children()) if listing.get_children() else None
            questions = self.db.favorite_list(sheet)
            title.config(text=f"{sheet} · 收藏题 · {len(questions)} 题")
            for q in questions:
                listing.insert("", "end", iid=q.qid, values=(q.sheet + " · " + q.stem.replace("\n", " "),))
            if questions:
                listing.selection_set(questions[0].qid)
            select()
        def unstar():
            if listing.selection():
                self.db.set_favorite(listing.selection()[0], False)
                refresh()
                if self.order:
                    q = self._q_at(self.pos)
                    self.btn_favorite.config(text="★ 已收藏 · 取消" if self.db.is_favorite(q.qid) else "☆ 收藏题目")
        remove = ttk.Button(body, text="取消收藏", command=unstar)
        remove.grid(row=2, column=1, sticky="w", padx=16, pady=(12, 0))
        listing.bind("<<TreeviewSelect>>", select)
        refresh()

    def show_bank(self):
        """回到「我的题库」导航页（当前做题进度已自动保存）"""
        self._cancel_auto_next()
        self._save_session()
        self._set_view("bank")

    # ---------------- 导航页渲染 ----------------

    def _sheet_progress(self, sheet: str):
        """返回某题库的进度信息（基于本会话保存的 sheet_sessions）"""
        saved = self._sheet_sessions.get(sheet)
        if not saved or not saved.get("order"):
            return None
        order = saved["order"]
        answers = saved.get("answers", {})
        done = sum(1 for qid in order if qid in answers)
        correct = sum(1 for qid in order
                      if answers.get(qid, {}).get("ok"))
        return {
            "total": len(order),
            "done": done,
            "correct": correct,
            "pos": saved.get("pos", 0),
            "finished": done >= len(order),
        }

    def _render_bank(self):
        """重建「我的题库」导航页"""
        for w in self.bank_cards:
            w.destroy()
        self.bank_cards = []

        total = sum(len(v) for v in self.sheets.values())
        if not self.sheet_names:
            self.bank_subtitle.config(text="还没有题库，点击下方「导入题库」开始")
        else:
            self.bank_subtitle.config(
                text=f"{len(self.sheet_names)} 套试卷 · 共 {total} 题 · 每套题库独立保存进度、错题和收藏")

        # 1) 各试卷卡片
        for sheet in self.sheet_names:
            qs = self.sheets.get(sheet, [])
            n = len(qs)
            # 题型分布
            by_type = {"单选": 0, "多选": 0, "判断": 0}
            for q in qs:
                by_type[q.type] = by_type.get(q.type, 0) + 1
            parts = [f"{t} {c}" for t, c in by_type.items() if c]
            prog = self._sheet_progress(sheet)
            wrong_n = len(self.db.wrong_list(sheet))
            favorite_n = len(self.db.favorite_list(sheet))

            card = tk.Frame(self.bank_body, bg=BG_CARD, highlightthickness=1,
                            highlightbackground=BORDER)
            self.bank_cards.append(card)
            card.pack(fill="x", pady=(0, 12))

            # 卡片内：左信息 + 右按钮
            info = tk.Frame(card, bg=BG_CARD)
            info.pack(side="left", fill="both", expand=True, padx=18, pady=16)
            # 试卷名 + 题数
            tk.Label(info, text=sheet, bg=BG_CARD, fg=FG_TEXT,
                     font=_font(12, "bold"), anchor="w",
                     wraplength=520, justify="left").pack(anchor="w")
            tk.Label(info, text=f"共 {n} 题   ·   " + " / ".join(parts),
                     bg=BG_CARD, fg=FG_MUTED, font=_font(10),
                     anchor="w").pack(anchor="w", pady=(4, 0))
            tk.Label(info, text=f"错题 {wrong_n}　·　收藏 {favorite_n}",
                     bg=BG_CARD, fg=FG_MUTED, font=_font(10),
                     anchor="w").pack(anchor="w", pady=(4, 0))

            # 进度条
            if prog:
                pct = prog["done"] / prog["total"] * 100 if prog["total"] else 0
                acc = (prog["correct"] / prog["done"] * 100) if prog["done"] else 0
                prog_lbl = "已完成" if prog["finished"] else "进行中"
                tk.Label(info,
                         text=f"{prog_lbl}　{prog['done']} / {prog['total']} 题 · 正确 {prog['correct']} · 准确率 {acc:.0f}%",
                         bg=BG_CARD, fg=(C_OK if prog["finished"] else ACCENT),
                         font=_font(10, "bold"), anchor="w").pack(anchor="w", pady=(8, 2))
                bar = ttk.Progressbar(info, length=300, maximum=100,
                                      value=pct, orient="horizontal")
                bar.pack(anchor="w")
            else:
                tk.Label(info, text="尚未开始", bg=BG_CARD, fg=FG_FAINT,
                         font=_font(10), anchor="w").pack(anchor="w", pady=(8, 0))

            # 右侧按钮：继续 / 开始 + 删除
            # 固定按钮列宽度，保证所有卡片右按钮列对齐（否则单按钮卡片
            # 的按钮会按文字宽度自适应，比双按钮卡片的列宽更宽，右缘不齐）
            btn_frame = tk.Frame(card, bg=BG_CARD)
            btn_frame.pack(side="right", fill="y", padx=(0, 18), pady=16)
            label = "继续刷题" if prog and not prog["finished"] else \
                    ("再练一遍" if prog and prog["finished"] else "开始刷题")
            ttk.Button(btn_frame, text=label, style="Accent.TButton",
                       width=23,
                       command=lambda s=sheet: self._open_sheet(s)).pack(
                           fill="x", pady=(0, 6))

            collection_row = tk.Frame(btn_frame, bg=BG_CARD)
            collection_row.pack(fill="x", pady=(0, 6))
            ttk.Button(collection_row, text=f"错题 {wrong_n}", width=10,
                       style="Soft.TButton",
                       command=lambda s=sheet: self._open_wrong(s)).pack(
                           side="left", fill="x", expand=True, padx=(0, 3))
            ttk.Button(collection_row, text=f"收藏 {favorite_n}", width=10,
                       style="Soft.TButton",
                       command=lambda s=sheet: self.show_favorites(s)).pack(
                           side="left", fill="x", expand=True, padx=(3, 0))
            # 删除按钮（红色危险色，小一号）；撑满按钮列宽，与上方主按钮左右对齐
            
            del_btn = ttk.Button(btn_frame, text="删除", style="Danger.TButton",
                                 command=lambda s=sheet: self._delete_sheet(s))
            del_btn.pack(fill="x")

        # 2) 导入卡片
        imp = tk.Frame(self.bank_body, bg=BG_CARD, highlightthickness=1,
                       highlightbackground=BORDER)
        self.bank_cards.append(imp)
        imp.pack(fill="x", pady=(4, 12))
        imp_info = tk.Frame(imp, bg=BG_CARD)
        imp_info.pack(side="left", fill="both", expand=True, padx=18, pady=16)
        tk.Label(imp_info, text="导入新题库", bg=BG_CARD, fg=FG_TEXT,
                 font=_font(12, "bold"), anchor="w").pack(anchor="w")
        tk.Label(imp_info,
                 text="从 Excel（.xlsx / .xls）导入，每个 sheet 是一套试卷\n支持单选 / 多选 / 判断，导入后永久保存",
                 bg=BG_CARD, fg=FG_MUTED, font=_font(10),
                 anchor="w", justify="left").pack(anchor="w", pady=(4, 0))
        imp_btn = tk.Frame(imp, bg=BG_CARD, width=120)
        imp_btn.pack(side="right", fill="y", padx=(0, 18), pady=16)
        imp_btn.pack_propagate(False)
        ttk.Button(imp_btn, text="导入题库", style="Accent.TButton",
                   command=self.on_import).pack(anchor="center")

        # 2b) 智能筛题卡片
        flt = tk.Frame(self.bank_body, bg=BG_CARD, highlightthickness=1,
                       highlightbackground=BORDER)
        self.bank_cards.append(flt)
        # 智能筛题入口暂时隐藏，保留原功能实现。
        flt_info = tk.Frame(flt, bg=BG_CARD)
        flt_info.pack(side="left", fill="both", expand=True, padx=18, pady=16)
        tk.Label(flt_info, text="智能筛题", bg=BG_CARD, fg=FG_TEXT,
                 font=_font(12, "bold"), anchor="w").pack(anchor="w")
        tk.Label(flt_info,
                 text="按背题规则筛选：判断错题、多选错误选项全含关键词、单选唯一最长正确答案的题剔除\n"
                      "剩下的存为新试卷「背题库」，方便集中背诵",
                 bg=BG_CARD, fg=FG_MUTED, font=_font(10),
                 anchor="w", justify="left").pack(anchor="w", pady=(4, 0))
        flt_btn = tk.Frame(flt, bg=BG_CARD, width=120)
        flt_btn.pack(side="right", fill="y", padx=(0, 18), pady=16)
        flt_btn.pack_propagate(False)
        ttk.Button(flt_btn, text="智能筛题",
                   style="Accent.TButton" if self.sheet_names else "Soft.TButton",
                   command=self.on_smart_filter).pack(anchor="center")

        # 重建卡片后，把滚轮/中键滚动手势绑到 bank_body 及所有后代控件，
        # 否则鼠标停在卡片上时滚轮事件被内层 widget 吃掉，导致无法滚动
        self.root.update_idletasks()
        self._bind_bank_wheel(self.bank_body)
        # 关键：把内嵌窗口宽度锁定为 canvas 宽度，防止卡片按内容自然宽度
        # 展开（曾出现 1352px > 视口，把右侧按钮推出屏幕外）。
        cw = self.bank_canvas.winfo_width()
        if cw > 1:
            self.bank_canvas.itemconfig(self._bank_win, width=cw)
        # 确保 scrollregion 覆盖全部卡片
        self.bank_canvas.config(scrollregion=self.bank_canvas.bbox("all"))

    def _open_sheet(self, sheet: str):
        """从导航页打开某套试卷（有进度则续接，否则重新开始）"""
        self._set_view("quiz")
        saved = self._sheet_sessions.get(sheet)
        valid_ids = {q.qid for q in self.sheets[sheet]}
        if saved and any(qid in valid_ids for qid in saved.get("order", [])):
            self._restore_session(copy.deepcopy(saved))
        else:
            self.start_session(sheet)

    def _delete_sheet(self, sheet: str):
        """删除某套试卷（含该卷的错题本记录与刷题进度）"""
        n = len(self.sheets.get(sheet, []))
        msg = (f"确定删除「{sheet}」？\n\n"
               f"该卷共 {n} 题将被永久移除，\n同时清除其错题本记录与刷题进度。")
        if not self.dialogs.askyesno("删除题库", msg, icon="warning"):
            return
        self.db.remove_sheet(sheet)
        # 清理内存与保存的进度
        self.sheets.pop(sheet, None)
        if sheet in self.sheet_names:
            self.sheet_names.remove(sheet)
        self._sheet_sessions.pop(sheet, None)
        self._wrong_sessions.pop(sheet, None)
        # 刷新界面
        self._load_from_db()
        self._save_session()
        self._render_bank()
        self.status_label.config(text=f"已删除「{sheet}」")

    def _delete_all(self):
        """清空整个题库（所有试卷 + 错题本 + 全部进度）"""
        if not self.sheet_names:
            self.dialogs.showinfo("提示", "题库为空，无需清空。")
            return
        n_sheets = len(self.sheet_names)
        n_q = sum(len(v) for v in self.sheets.values())
        msg = (f"确定清空全部题库？\n\n"
               f"将永久删除全部 {n_sheets} 套试卷、{n_q} 道题，\n"
               f"以及错题本和所有刷题进度。\n此操作不可撤销。")
        if not self.dialogs.askyesno("清空题库", msg, icon="warning"):
            return
        self.db.remove_all()
        self.sheets = {}
        self.sheet_names = []
        self._sheet_sessions = {}
        self._wrong_sessions = {}
        self._load_from_db()
        self._save_session()
        self._render_bank()
        self.status_label.config(text="已清空全部题库")

    def _open_wrong(self, sheet: str):
        """从某套题库卡片进入该题库的错题重做。"""
        if not self.db.wrong_list(sheet):
            self.dialogs.showinfo("提示", f"「{sheet}」暂无错题，继续加油！")
            return
        self._set_view("quiz")
        saved = self._wrong_sessions.get(sheet)
        if saved:
            self._normal_session = copy.deepcopy(self._sheet_sessions.get(sheet))
            self._restore_session(copy.deepcopy(saved))
        else:
            self.start_session(sheet, wrong_only=True)
        self.status_label.config(text=f"{sheet} · 错题重做")

    def _refresh_order_box(self):
        self.order_combo.config(state="readonly"
                                if (self.mode != "wrong" and self.order)
                                else "disabled")

    def _order_mode(self) -> str:
        """下拉框值 → 内部模式 random / original / type"""
        v = self.order_var.get()
        if v == "原顺序":
            return "original"
        if v == "题型分组":
            return "type"
        return "random"

    def _set_order_display(self, mode: str):
        self.order_var.set({
            "original": "原顺序",
            "type": "题型分组",
            "random": "随机乱序",
        }.get(mode, "随机乱序"))

    def _build_order(self, qs: list[Question]) -> tuple[list[str], str]:
        """按当前顺序模式生成本轮 qid 顺序。
        - original：导入顺序
        - random：随机乱序
        - type：按 单选→多选→判断 分组，组内保持原顺序
        """
        qids = [q.qid for q in qs]
        mode = self._order_mode()
        if mode == "original":
            return qids, "original"
        if mode == "type":
            rank = {"单选": 0, "多选": 1, "判断": 2}
            # 稳定排序：同题型保持导入顺序
            ordered = sorted(qs, key=lambda q: rank.get(q.type, 3))
            return [q.qid for q in ordered], "type"
        # random
        out = list(qids)
        random.shuffle(out)
        return out, "random"

    # ---------------- 数据库加载 ----------------

    def _load_from_db(self):
        self.sheet_names = self.db.sheet_names()
        self.sheets = {s: self.db.get_sheet(s) for s in self.sheet_names}
        self.sheet_combo["values"] = self.sheet_names
        if self.sheet_names:
            self.sheet_combo.config(state="readonly")
       # self._update_header()

    # ---------------- 导入 ----------------

    def on_import(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 题库",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            sheets, skipped = load_excel(path)
        except ValueError as e:
            self.dialogs.showerror("导入失败", str(e))
            return
        if not sheets:
            msg = "未从文件中读取到任何题目。\n请确认表头包含「题目」列，且每行都有题目和答案。"
            if skipped:
                msg += "\n\n" + "\n".join("· " + s for s in skipped[:20])
                if len(skipped) > 20:
                    msg += f"\n· … 共 {len(skipped)} 条提示"
            self.dialogs.showerror("导入失败", msg)
            return
        # 写入长期数据库
        source = os.path.basename(path)
        for i, (sheet, qs) in enumerate(sheets.items()):
            self.db.upsert_sheet(sheet, source, qs, seq=i)
        self._load_from_db()

        # 导入后停留在导航页，刷新各卷卡片；用户点「开始刷题」进入
        self._set_view("bank")
        self._save_session()

        total = sum(len(v) for v in self.sheets.values())
        msg = f"已导入并存档「{source}」，共 {total} 题"
        if skipped:
            msg += f"（{len(skipped)} 行未识别，已跳过）"
            try:
                log_path = os.path.join(os.path.dirname(path) or ".",
                                        os.path.basename(path) + ".import_log.txt")
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write("导入时间：" + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
                    f.write(f"题库：{source}（读取 {total} 题，跳过 {len(skipped)} 行）\n\n")
                    for s in skipped:
                        f.write("- " + s + "\n")
                msg += f"（详情见 {os.path.basename(log_path)}）"
            except OSError:
                pass
        self._toast(msg)

    # ---------------- 智能筛题（背题规则） ----------------

    def on_smart_filter(self):
        """按背题规则筛选当前题库：剔除"不用背"的题，剩下的存为新试卷「背题库」。

        规则见 memorization_filter.py。筛选只读，不改动原试卷；
        结果作为一套独立试卷写入题库，可继续刷题 / 删除。
        """
        if not self.sheet_names:
            self.dialogs.showinfo("提示", "还没有题库，请先导入 Excel 题库。")
            return
        all_qs: list[Question] = []
        for sheet in self.sheet_names:
            all_qs.extend(self.sheets.get(sheet, []))
        if not all_qs:
            self.dialogs.showinfo("提示", "题库里没有可筛选的题目。")
            return

        kept, removed = filter_questions(all_qs)
        if not kept:
            self.dialogs.showinfo("智能筛题",
                                "按规则筛选后没有剩余题目（全部被剔除）。")
            return
        if len(kept) == len(all_qs):
            self.dialogs.showinfo("智能筛题",
                                "没有题目符合剔除规则，整份题库都建议保留。")
            return

        # 汇总剔除原因（按题型）
        drop_by_type = {"单选": 0, "多选": 0, "判断": 0}
        for q, _ in removed:
            drop_by_type[q.type] = drop_by_type.get(q.type, 0) + 1
        detail = "、".join(f"{t} {c}" for t, c in drop_by_type.items() if c)

        # 作为新试卷写入（复用源文件作为来源标记；行号重排 1..N）
        new_sheet = "背题库"
        new_qs = []
        for i, q in enumerate(kept, start=1):
            new_qs.append(Question(
                qid=f"{new_sheet}__{i}", sheet=new_sheet, row=i,
                type=q.type, stem=q.stem, options=list(q.options),
                answer=list(q.answer), explanation=q.explanation,
                difficulty=q.difficulty,
            ))
        seq = len(self.sheet_names)
        self.db.upsert_sheet(new_sheet, "智能筛题", new_qs, seq=seq)
        self._load_from_db()
        self._save_session()
        self._set_view("bank")
        self._render_bank()

        msg = (f"已生成「{new_sheet}」：{len(kept)} 题待背\n"
               f"剔除 {len(removed)} 题（{detail}）\n"
               f"规则：判断错题/多选错误选项全含关键词/单选唯一最长正确答案")
        self._toast("已生成背题库")
        # 用信息框给出更完整的说明（不含 defaultbutton，兼容当前 Tk 构建）
        self.dialogs.showinfo("智能筛题", msg)

    # ---------------- 综合设置（外观 + 答题） ----------------

    def _init_settings(self, path):
        """加载 app_settings.json，把字体缩放与整套配色落到全局常量。"""
        self._settings_path = path
        cfg = settings_mod.load_config(path)
        self._settings_cfg = cfg
        global FONT_SCALE
        FONT_SCALE = settings_mod.font_scale_factor(cfg["font_scale"])
        self._theme_mode = cfg["theme"]
        self._resolved_theme = None
        self._settings_window = None
        self._theme_watch_id = None
        self._apply_color_theme(self._theme_mode, refresh_widgets=False)
        self._auto_next_delay = cfg["auto_next_delay"]
        self._default_order = cfg["default_order"]

    def _apply_color_theme(self, mode, refresh_widgets=True, accent_id=None):
        """像编辑器主题一样切换背景、表面、文字、边框和控件配色。"""
        global BG_APP, BG_CARD, BG_SUBTLE, BORDER, BORDER_SOFT
        global FG_TEXT, FG_MUTED, FG_FAINT, FG_ON_ACCENT
        global ACCENT, ACCENT_HOVER, ACCENT_ACTIVE, ACCENT_SOFT
        global C_UNANSWERED, C_TEXT_ON_LIGHT, C_OK_SOFT, C_BAD_SOFT, C_CURRENT
        resolved = settings_mod.resolve_theme(mode)
        palette = dict(settings_mod.THEME_PALETTES[resolved])
        # 主题色：优先用指定 id，否则用当前配置
        if accent_id is None:
            accent_id = getattr(self, "_settings_cfg", {}).get("accent", "midnight")
        palette.update(settings_mod.accent_colors(accent_id))
        names = ("BG_APP", "BG_CARD", "BG_SUBTLE", "BORDER", "BORDER_SOFT",
                 "FG_TEXT", "FG_MUTED", "FG_FAINT", "FG_ON_ACCENT", "ACCENT",
                 "ACCENT_HOVER", "ACCENT_ACTIVE", "ACCENT_SOFT", "C_UNANSWERED",
                 "C_TEXT_ON_LIGHT", "C_OK_SOFT", "C_BAD_SOFT")
        old = {name: globals()[name] for name in names}
        for name in names:
            globals()[name] = palette[name]
        C_CURRENT = ACCENT
        self._theme_mode = mode
        self._resolved_theme = resolved
        self._reconfigure_styles()
        if refresh_widgets and hasattr(self, "root"):
            background_names = ("BG_APP", "BG_CARD", "BG_SUBTLE", "BORDER", "BORDER_SOFT",
                                "ACCENT", "ACCENT_HOVER", "ACCENT_ACTIVE", "ACCENT_SOFT",
                                "C_UNANSWERED", "C_OK_SOFT", "C_BAD_SOFT")
            foreground_names = ("FG_TEXT", "FG_MUTED", "FG_FAINT", "FG_ON_ACCENT",
                                "C_TEXT_ON_LIGHT", "ACCENT")
            bg_map = {old[name].lower(): palette[name] for name in background_names}
            fg_map = {old[name].lower(): palette[name] for name in foreground_names}
            self.root.configure(bg=BG_APP)
            self._recolor_widget_tree(self.root, bg_map, fg_map)
            if hasattr(self, "_paint_matrix"):
                self._paint_matrix()
            self.root.update_idletasks()

    def _recolor_widget_tree(self, widget, bg_map, fg_map):
        """只替换旧主题色，保留题目对错等语义色和控件自定义状态。"""
        background_options = ("background", "activebackground", "highlightbackground",
                              "highlightcolor", "selectbackground", "troughcolor")
        foreground_options = ("foreground", "activeforeground", "selectforeground",
                              "insertbackground")
        try:
            available = widget.configure()
            changes = {}
            for option in background_options + foreground_options:
                if option not in available:
                    continue
                value = str(widget.cget(option)).lower()
                color_map = bg_map if option in background_options else fg_map
                if value in color_map:
                    changes[option] = color_map[value]
            if changes:
                widget.configure(**changes)
        except tk.TclError:
            pass
        for child in widget.winfo_children():
            self._recolor_widget_tree(child, bg_map, fg_map)

    def _watch_system_theme(self):
        """跟随 Windows 应用主题变化，行为接近 VS Code 的自动主题。"""
        if self._theme_mode == "system":
            resolved = settings_mod.resolve_theme("system")
            if resolved != self._resolved_theme:
                self._apply_color_theme("system")
        self._theme_watch_id = self.root.after(2000, self._watch_system_theme)

    def _save_settings(self):
        self._settings_cfg = settings_mod.save_config(
            self._settings_path, self._settings_cfg)

    def _reconfigure_styles(self):
        """重设 ttk 样式（依赖全局颜色常量，调用后需 root.update 生效）。"""
        style = ttk.Style()
        style.configure("TFrame", background=BG_CARD)
        style.configure("TLabel", background=BG_CARD, foreground=FG_TEXT, font=_font(10))
        style.configure("Treeview", font=_font(11), rowheight=40, background=BG_CARD,
                        fieldbackground=BG_CARD, foreground=FG_TEXT, borderwidth=0)
        style.configure("Treeview.Heading", font=_font(10, "bold"), background=BG_SUBTLE,
                        foreground=FG_TEXT)
        # --- 通用按钮 ---
        style.configure("TButton", font=_font(10), padding=(12, 7), borderwidth=0,
                        relief="flat", background=BG_CARD, foreground=FG_TEXT,
                        lightcolor=BG_CARD, darkcolor=BG_CARD,
                        focuscolor=ACCENT_SOFT, focusthickness=0)
        style.map("TButton",
                  background=[("active", BG_SUBTLE), ("disabled", BG_SUBTLE)],
                  foreground=[("disabled", FG_FAINT)])
        # 主题按钮
        style.configure("Accent.TButton", font=_font(10, "bold"), padding=(14, 8),
                        borderwidth=0, relief="flat", background=ACCENT,
                        foreground=FG_ON_ACCENT, lightcolor=ACCENT, darkcolor=ACCENT,
                        focuscolor=ACCENT_SOFT, focusthickness=0)
        style.map("Accent.TButton",
                  background=[("active", ACCENT_HOVER), ("pressed", ACCENT_ACTIVE),
                             ("disabled", BG_SUBTLE)],
                  foreground=[("disabled", FG_FAINT)])
        # 次级/幽灵按钮
        style.configure("Soft.TButton", font=_font(10), padding=(12, 7),
                        borderwidth=1, relief="flat", background=BG_CARD,
                        foreground=FG_TEXT, lightcolor=BORDER, darkcolor=BORDER,
                        focuscolor=ACCENT_SOFT, focusthickness=0)
        style.map("Soft.TButton", background=[("active", ACCENT_SOFT), ("disabled", BG_SUBTLE)],
                  foreground=[("disabled", FG_FAINT)])
        # 删除按钮（灰色弱化）
        style.configure("Danger.TButton", font=_font(9), padding=(10, 5),
                        borderwidth=1, relief="flat", background=BG_CARD,
                        foreground=FG_MUTED, lightcolor=BORDER, darkcolor=BORDER,
                        focuscolor=BG_SUBTLE, focusthickness=0)
        style.map("Danger.TButton",
                  background=[("active", BG_SUBTLE), ("disabled", BG_SUBTLE)],
                  foreground=[("active", FG_MUTED), ("disabled", FG_FAINT)])
        # Treeview 选中色
        style.map("Treeview", background=[("selected", ACCENT_SOFT)], foreground=[("selected", FG_TEXT)])
        style.configure("Header.TLabel", font=_font(18, "bold"), background=BG_APP, foreground=FG_TEXT)
        style.configure("Sub.TLabel", font=_font(10), background=BG_APP, foreground=FG_MUTED)
        style.configure("SubCard.TLabel", font=_font(10), background=BG_SUBTLE, foreground=FG_MUTED)
        style.configure("Question.TLabel", font=_font(16), background=BG_CARD, foreground=FG_TEXT)
        style.configure("Explain.TLabel", font=_font(11), background=BG_SUBTLE, foreground=FG_MUTED)
        style.configure("Ok.TLabel", font=_font(11, "bold"), background=C_OK_SOFT, foreground=C_OK)
        style.configure("Bad.TLabel", font=_font(11, "bold"), background=C_BAD_SOFT, foreground=C_BAD)
        style.configure("Option.TCheckbutton", font=_font(11), background=BG_CARD, foreground=FG_TEXT)
        style.configure("TCombobox", fieldbackground=BG_CARD, background=BG_CARD,
                        foreground=FG_TEXT, arrowcolor=ACCENT, bordercolor=BORDER,
                        lightcolor=BG_CARD, darkcolor=BG_CARD, selectbackground=ACCENT,
                        selectforeground=FG_ON_ACCENT, padding=5)
        style.map("TCombobox", fieldbackground=[("readonly", BG_CARD), ("disabled", BG_SUBTLE)],
                  foreground=[("readonly", FG_TEXT), ("disabled", FG_FAINT)],
                  arrowcolor=[("readonly", ACCENT)], bordercolor=[("focus", ACCENT)])
        style.configure("TProgressbar", background=ACCENT, troughcolor=BG_SUBTLE,
                        bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT)
        style.configure("Vertical.TScrollbar", background=BORDER, troughcolor=BG_APP,
                        bordercolor=BG_APP, arrowcolor=FG_MUTED, relief="flat")
        # 底部导航按钮
        style.configure("Nav.Soft.TButton", font=_font(10), padding=(16, 10))
        style.configure("Nav.Accent.TButton", font=_font(10, "bold"), padding=(16, 10))

    def open_settings(self):
        """打开唯一的综合设置窗口；重复点击只唤醒已有窗口。"""
        existing = self._settings_window
        try:
            if existing is not None and existing.winfo_exists():
                existing.deiconify()
                existing.lift()
                existing.focus_force()
                return
        except tk.TclError:
            pass
        win = tk.Toplevel(self.root)
        self._settings_window = win
        win.title("设置")
        win.transient(self.root)
        win.resizable(False, False)
        win.configure(bg=BG_APP)

        def close_settings():
            self._settings_window = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", close_settings)

        # 顶部标题
        header = tk.Frame(win, bg=BG_CARD, padx=24, pady=16)
        header.pack(fill="x", side="top")
        tk.Label(header, text="设置", bg=BG_CARD, fg=FG_TEXT,
                 font=_font(16, "bold")).pack(side="left")

        # 内容区：左导航 + 右面板
        content = tk.Frame(win, bg=BG_APP)
        content.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        content.columnconfigure(0, weight=0)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        # 左导航
        nav = tk.Frame(content, bg=BG_SUBTLE, width=150, highlightthickness=1,
                       highlightbackground=BORDER)
        nav.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        nav.pack_propagate(False)
        tk.Label(nav, text="设置项", bg=BG_SUBTLE, fg=FG_MUTED,
                 font=_font(10, "bold")).pack(anchor="w", padx=14, pady=(14, 8))

        nav_vars = {}
        sections = [("AI 模型", "模型接口配置"), ("外观", "字体大小 · 界面主题"),
                    ("答题", "自动跳转 · 默认顺序")]
        nav_btns = {}
        _show_ref = [None]  # 延迟绑定：show_section 定义后填入
        for i, (title, sub) in enumerate(sections):
            row = tk.Frame(nav, bg=BG_SUBTLE)
            row.pack(fill="x", padx=8, pady=2)
            row.columnconfigure(0, weight=1)
            title_label = tk.Label(row, text=title, bg=BG_SUBTLE, fg=FG_TEXT,
                                   font=_font(11, "bold"), cursor="hand2")
            title_label.grid(row=0, column=0, sticky="w", padx=8, pady=(8, 0))
            sub_label = tk.Label(row, text=sub, bg=BG_SUBTLE, fg=FG_MUTED,
                                 font=_font(9), cursor="hand2")
            sub_label.grid(row=1, column=0, sticky="w", padx=8, pady=(0, 8))
            row.configure(cursor="hand2")
            callback = lambda e, _t=title, _r=_show_ref: _r[0](_t)
            row.bind("<Button-1>", callback)
            title_label.bind("<Button-1>", callback)
            sub_label.bind("<Button-1>", callback)
            nav_btns[title] = row

        # 右面板容器
        panels = tk.Frame(content, bg=BG_CARD, highlightthickness=1,
                          highlightbackground=BORDER)
        panels.grid(row=0, column=1, sticky="nsew")
        panels.columnconfigure(0, weight=1)
        panels.rowconfigure(0, weight=1)

        # ---------- AI 模型面板 ----------
        url_entry_var = tk.StringVar()
        def _build_ai_panel(parent, _url_var=url_entry_var):
            f = tk.Frame(parent, bg=BG_CARD, padx=24, pady=20)
            f.grid(row=0, column=0, sticky="nsew")
            f.columnconfigure(1, weight=1)
            tk.Label(f, text="模型接口", bg=BG_CARD, fg=FG_TEXT,
                     font=_font(13, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
            tk.Label(f, text="OpenAI / LM Studio / vLLM 选「OpenAI 兼容」，Ollama 选「Ollama」。",
                     bg=BG_CARD, fg=FG_MUTED, font=_font(9), wraplength=520,
                     justify="left").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 12))
            cfg = self._ai_config
            variables = {}
            rows = [("provider", "接口协议"), ("base_url", "Base URL"),
                    ("model", "模型名称"), ("api_key", "API Key"),
                    ("timeout", "超时（秒）")]
            for i, (key, title) in enumerate(rows):
                r = i + 2
                tk.Label(f, text=title, bg=BG_CARD, fg=FG_MUTED,
                         font=_font(10)).grid(row=r, column=0, sticky="w", pady=7, padx=(0, 14))
                if key == "provider":
                    entry = ttk.Combobox(f, values=["OpenAI 兼容", "Ollama"],
                                         state="readonly", width=30)
                    entry.set(cfg.get("provider", "OpenAI 兼容"))

                    def change(event=None, _e=entry, _v=_url_var):
                        if _e.get() == "Ollama":
                            _v.set("http://localhost:11434")
                        else:
                            _v.set("https://api.openai.com/v1")
                    entry.bind("<<ComboboxSelected>>", change)
                    variables["provider"] = entry
                elif key == "base_url":
                    entry = ttk.Entry(f, textvariable=_url_var, width=40)
                    _url_var.set(cfg.get("base_url", ""))
                    variables["base_url"] = entry
                elif key == "api_key":
                    entry = ttk.Entry(f, width=40, show="*")
                    entry.insert(0, cfg.get("api_key", ""))
                    variables["api_key"] = entry
                else:
                    entry = ttk.Entry(f, width=16)
                    entry.insert(0, str(cfg.get(key, "")))
                    variables[key] = entry
                entry.grid(row=r, column=1, sticky="w")
            tk.Label(f, text="密钥永久保存在本配置文件中，也可用 OPENAI_API_KEY 环境变量提供（已保存的优先）。",
                     bg=BG_CARD, fg=FG_FAINT, font=_font(9), wraplength=520,
                     justify="left").grid(row=8, column=0, columnspan=2, sticky="w", pady=(14, 0))
            ttk.Button(f, text="保存模型设置", style="Accent.TButton",
                       command=lambda: self._save_ai_from_panel(variables)).grid(
                row=9, column=0, columnspan=2, sticky="e", pady=(18, 0))
            return f

        def _save_ai_from_panel(variables):
            config = {
                "provider": variables["provider"].get().strip(),
                "base_url": variables["base_url"].get().strip(),
                "model": variables["model"].get().strip(),
                "api_key": variables["api_key"].get().strip(),
                "timeout": variables["timeout"].get().strip(),
            }
            try:
                config["timeout"] = int(config["timeout"])
                if not 5 <= config["timeout"] <= 600:
                    raise ValueError("超时请设置为 5–600 秒。")
                from .ai_explanation import endpoint
                endpoint(config)
            except (ValueError, OSError) as exc:
                self.dialogs.showerror("无法保存", str(exc), parent=win)
                return
            try:
                with open(self._ai_path + ".tmp", "w", encoding="utf-8") as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                os.replace(self._ai_path + ".tmp", self._ai_path)
            except OSError as exc:
                self.dialogs.showerror("无法保存", str(exc), parent=win)
                return
            self._ai_config = config
            self.dialogs.showinfo("已保存", "模型设置已保存。")

        # ---------- 外观面板 ----------
        def _build_appearance_panel(parent):
            f = tk.Frame(parent, bg=BG_CARD, padx=24, pady=20)
            f.grid(row=0, column=0, sticky="nsew")
            f.columnconfigure(1, weight=1)
            tk.Label(f, text="外观", bg=BG_CARD, fg=FG_TEXT,
                     font=_font(13, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")

            # 字体大小
            tk.Label(f, text="字体大小", bg=BG_CARD, fg=FG_TEXT,
                     font=_font(11, "bold")).grid(row=1, column=0, columnspan=2, sticky="w", pady=(16, 4))
            font_frame = tk.Frame(f, bg=BG_CARD)
            font_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
            font_frame.columnconfigure(1, weight=1)
            font_var = tk.DoubleVar(value=FONT_SCALE)
            scale = tk.Scale(font_frame, from_=0.8, to=1.5, resolution=0.1,
                             orient="horizontal", variable=font_var,
                             bg=BG_CARD, troughcolor=BORDER,
                             font=_font(10), length=420, sliderrelief="flat",
                             highlightthickness=0, showvalue=False)
            scale.grid(row=0, column=0, sticky="w")
            font_val_lbl = tk.Label(font_frame, text="100%", bg=BG_CARD, fg=FG_TEXT,
                                    font=_font(10, "bold"))
            font_val_lbl.grid(row=0, column=1, sticky="e")
            font_var.trace_add("write", lambda *_a: font_val_lbl.config(
                text=f"{round(font_var.get() * 100)}%"))

            # 界面主题
            tk.Label(f, text="界面主题", bg=BG_CARD, fg=FG_TEXT,
                     font=_font(11, "bold")).grid(row=3, column=0, columnspan=2, sticky="w", pady=(20, 6))
            theme_frame = tk.Frame(f, bg=BG_CARD)
            theme_frame.grid(row=4, column=0, columnspan=2, sticky="ew")
            theme_cards = {}

            def refresh_theme_cards():
                for key, card in theme_cards.items():
                    selected = key == self._theme_mode
                    card.config(highlightbackground=ACCENT if selected else BORDER,
                                bg=ACCENT_SOFT if selected else BG_SUBTLE)
                    for child in card.winfo_children():
                        child.config(bg=ACCENT_SOFT if selected else BG_SUBTLE)

            def select_theme(mode):
                self._settings_cfg["theme"] = mode
                self._apply_color_theme(mode)
                self._save_settings()
                refresh_theme_cards()

            for i, mode in enumerate(("system", "light", "dark")):
                card = tk.Frame(theme_frame, bg=BG_SUBTLE, cursor="hand2",
                                highlightthickness=2, highlightbackground=BORDER,
                                padx=14, pady=12)
                card.grid(row=0, column=i, padx=(0, 10), sticky="ew")
                theme_frame.columnconfigure(i, weight=1)
                title = tk.Label(card, text=settings_mod.THEME_LABELS[mode], bg=BG_SUBTLE,
                                 fg=FG_TEXT, font=_font(10, "bold"), cursor="hand2")
                title.pack(anchor="w")
                hint = {"system": "随 Windows", "light": "明亮界面", "dark": "深色界面"}[mode]
                detail = tk.Label(card, text=hint, bg=BG_SUBTLE, fg=FG_MUTED,
                                  font=_font(9), cursor="hand2")
                detail.pack(anchor="w", pady=(3, 0))
                callback = lambda _e, _m=mode: select_theme(_m)
                card.bind("<Button-1>", callback)
                title.bind("<Button-1>", callback)
                detail.bind("<Button-1>", callback)
                theme_cards[mode] = card
            refresh_theme_cards()
            tk.Label(f, text="主题会同时切换背景、文字、边框和控件；跟随系统会自动响应 Windows 设置。",
                     bg=BG_CARD, fg=FG_FAINT, font=_font(9)).grid(
                row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))

            # ---------- 主题色 ----------
            tk.Label(f, text="主题色", bg=BG_CARD, fg=FG_TEXT,
                     font=_font(11, "bold")).grid(row=6, column=0, columnspan=2, sticky="w", pady=(20, 6))
            tk.Label(f, text="按钮、选中行、进度条使用此颜色。预设参考 Windows 11 与 VS Code。",
                     bg=BG_CARD, fg=FG_MUTED, font=_font(9)).grid(
                row=7, column=0, columnspan=2, sticky="w")
            accent_frame = tk.Frame(f, bg=BG_CARD)
            accent_frame.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(6, 0))
            accent_swatch_btns = {}

            def refresh_accent_swatches():
                current = self._settings_cfg.get("accent", "midnight")
                for aid, btn in accent_swatch_btns.items():
                    if aid == current:
                        btn.config(highlightbackground=ACCENT, highlightthickness=2)
                    else:
                        btn.config(highlightbackground=BORDER, highlightthickness=1)

            def select_accent(aid):
                self._settings_cfg["accent"] = aid
                self._apply_color_theme(self._theme_mode, accent_id=aid)
                self._save_settings()
                refresh_accent_swatches()

            for i, (aid, theme) in enumerate(settings_mod.ACCENT_THEMES.items()):
                btn = tk.Label(accent_frame, text=theme["label"],
                               bg=theme["ACCENT"], fg=theme["FG_ON_ACCENT"],
                               font=_font(9, "bold"), width=7, pady=6,
                               cursor="hand2", highlightthickness=1,
                               highlightbackground=BORDER)
                btn.grid(row=0, column=i, padx=3, sticky="ew")
                accent_frame.columnconfigure(i, weight=1)
                for w in (btn,):
                    w.bind("<Button-1>", lambda e, _a=aid: select_accent(_a))
                accent_swatch_btns[aid] = btn
            refresh_accent_swatches()

            # 保存外观
            def save_appearance():
                global FONT_SCALE
                fs = settings_mod.font_scale_factor(font_var.get())
                FONT_SCALE = fs
                self._settings_cfg["font_scale"] = fs
                self._save_settings()
                # 重刷样式
                self._reconfigure_styles()
                win.update_idletasks()
                self.dialogs.showinfo("已保存", "外观设置已保存，即时生效。")

            ttk.Button(f, text="保存外观设置", style="Accent.TButton",
                       command=save_appearance).grid(
                row=9, column=0, columnspan=2, sticky="e", pady=(20, 0))
            return f

        # ---------- 答题面板 ----------
        def _build_quiz_panel(parent):
            f = tk.Frame(parent, bg=BG_CARD, padx=24, pady=20)
            f.grid(row=0, column=0, sticky="nsew")
            f.columnconfigure(1, weight=1)
            tk.Label(f, text="答题", bg=BG_CARD, fg=FG_TEXT,
                     font=_font(13, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")

            # 自动跳转延迟
            tk.Label(f, text="答对后自动下一题延迟", bg=BG_CARD, fg=FG_TEXT,
                     font=_font(11, "bold")).grid(row=1, column=0, columnspan=2, sticky="w", pady=(16, 4))
            delay_frame = tk.Frame(f, bg=BG_CARD)
            delay_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
            delay_frame.columnconfigure(1, weight=1)
            delay_var = tk.IntVar(value=self._auto_next_delay)
            d_scale = tk.Scale(delay_frame, from_=100, to=3000, resolution=100,
                               orient="horizontal", variable=delay_var,
                               bg=BG_CARD, troughcolor=BORDER,
                               font=_font(10), length=420, sliderrelief="flat",
                               highlightthickness=0, showvalue=False)
            d_scale.grid(row=0, column=0, sticky="w")
            delay_val_lbl = tk.Label(delay_frame, text="0.7 秒", bg=BG_CARD, fg=FG_TEXT,
                                     font=_font(10, "bold"))
            delay_val_lbl.grid(row=0, column=1, sticky="e")
            delay_var.trace_add("write", lambda *_a: delay_val_lbl.config(
                text=f"{delay_var.get() / 1000:.1f} 秒"))

            # 默认顺序
            tk.Label(f, text="默认题目顺序", bg=BG_CARD, fg=FG_TEXT,
                     font=_font(11, "bold")).grid(row=3, column=0, columnspan=2, sticky="w", pady=(20, 4))
            order_var = tk.StringVar(value=self._default_order)
            order_combo = ttk.Combobox(f, textvariable=order_var, state="readonly",
                                       width=16,
                                       values=["随机乱序", "原顺序", "题型分组"])
            order_combo.set(settings_mod.ORDER_LABELS.get(self._default_order, "随机乱序"))
            order_combo.grid(row=4, column=0, columnspan=2, sticky="w")
            tk.Label(f, text="新建一轮答题时的初始顺序，仍可在顶栏随时切换。",
                     bg=BG_CARD, fg=FG_FAINT, font=_font(9)).grid(
                row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))

            # 保存答题设置
            def save_quiz():
                label_to_key = {v: k for k, v in settings_mod.ORDER_LABELS.items()}
                delay = int(settings_mod.clamp(delay_var.get(), 100, 3000))
                self._settings_cfg["auto_next_delay"] = delay
                self._settings_cfg["default_order"] = label_to_key.get(
                    order_combo.get(), "random")
                self._auto_next_delay = delay
                self._default_order = self._settings_cfg["default_order"]
                self._save_settings()
                self.dialogs.showinfo("已保存", "答题设置已保存。")

            ttk.Button(f, text="保存答题设置", style="Accent.TButton",
                       command=save_quiz).grid(
                row=6, column=0, columnspan=2, sticky="e", pady=(20, 0))
            return f

        panel_cache = {}
        panel_builders = {"AI 模型": _build_ai_panel, "外观": _build_appearance_panel,
                          "答题": _build_quiz_panel}

        def show_section(name):
            for child in panel_cache.values():
                child.grid_remove()
            for title, row in nav_btns.items():
                if title == name:
                    row.config(bg=ACCENT_SOFT)
                    for lbl in row.winfo_children():
                        lbl.config(bg=ACCENT_SOFT)
                else:
                    row.config(bg=BG_SUBTLE)
                    for lbl in row.winfo_children():
                        lbl.config(bg=BG_SUBTLE)
            if name not in panel_cache:
                panel_cache[name] = panel_builders[name](panels)
            panel_cache[name].grid()

        _show_ref[0] = show_section  # 延迟绑定：导航按钮点击 → 切换面板

        # 快捷入口
        footer = tk.Frame(win, bg=BG_CARD, padx=24, pady=12)
        footer.pack(fill="x", side="bottom")
        ttk.Button(footer, text="答题快捷键设置", style="Soft.TButton",
                   command=self.open_shortcut_settings).pack(side="left")
        tk.Button(footer, text="关闭", bg=BG_CARD, fg=FG_MUTED,
                  font=_font(10), relief="flat", cursor="hand2",
                  activebackground=BG_SUBTLE, activeforeground=FG_MUTED,
                  command=close_settings).pack(side="right")

        # 显示第一节
        show_section("AI 模型")
        # 让窗口自适应
        win.update_idletasks()
        w = max(win.winfo_reqwidth(), 620)
        h = max(win.winfo_reqheight(), 460)
        x = self.root.winfo_x() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.lift()
        win.focus_force()

    # ---------------- 会话管理 ----------------

    def start_session(self, sheet: str, wrong_only: bool = False):
        """开始一轮答题；顺序取顶栏选择（random 随机 / original 导入顺序）"""
        qs = self.db.wrong_list(sheet) if wrong_only else self.sheets.get(sheet, [])
        if not qs:
            self.dialogs.showinfo("提示", "没有可作答的题目。")
            return
        self._cancel_auto_next()
        if (wrong_only and self.mode == "normal" and self.order
                and self.current_sheet == sheet):
            self._normal_session = self._session_data()
        elif wrong_only:
            self._normal_session = copy.deepcopy(self._sheet_sessions.get(sheet))
        elif not wrong_only:
            self._normal_session = None
        if wrong_only:
            qs = self.db.wrong_list(sheet)
            self.current_sheet = sheet
            self.mode = "wrong"
        else:
            qs = self.sheets.get(sheet, [])
            self.current_sheet = sheet
            self.mode = "normal"
        if not qs:
            self.dialogs.showinfo("提示", "没有可作答的题目。")
            return
        self.placeholder.grid_remove()
        self.current_qs = list(qs)
        self.qmap = {q.qid: q for q in self.current_qs}
        if wrong_only:
            # 错题重做固定用导入顺序
            self.order = [q.qid for q in self.current_qs]
            self._round_order = "original"
        else:
            self.order, self._round_order = self._build_order(self.current_qs)
        self.pos = 0
        self.answers = {}
        self.score_correct = 0
        self.score_wrong = 0
        self._summary_mode = False
        self.sheet_var.set(sheet if not wrong_only else f"❌ {sheet} · 错题重做")
        if wrong_only:
            self.sheet_combo.config(state="disabled")
        self._set_view("quiz")
        self._refresh_navigation()
        self._refresh_order_box()
        self._render()
        self._save_session()

    def _on_order_change(self, event=None):
        """切换顺序 → 重新洗牌本轮（错题重做模式下不可切换）"""
        if self.mode == "wrong" or not self.order:
            return
        if not self.dialogs.askyesno("切换顺序",
                                   "切换题目顺序将重新开始本轮（已答题目不保留），确定？"):
            # 取消：还原下拉框到当前轮次的顺序
            self._set_order_display(getattr(self, "_round_order", "random"))
            return
        self.start_session(self.current_sheet)

    def _load_sessions_only(self):
        """启动时只加载会话进度（用于导航页显示各卷进度），不进入做题页。"""
        sess = load_session()
        if not sess:
            return
        self._sheet_sessions = sess.get("sheet_sessions", {})
        self._wrong_sessions = sess.get("wrong_sessions", {})
        self._normal_session = sess.get("normal_session")

    def _restore_session(self, sess=None):
        """启动时恢复上次的刷题位置"""
        from_disk = sess is None
        sess = load_session() if from_disk else sess
        if not sess:
            return
        if from_disk:
            self._sheet_sessions = sess.get("sheet_sessions", {})
            self._wrong_sessions = sess.get("wrong_sessions", {})
        if "normal_session" in sess:
            self._normal_session = sess.get("normal_session")
        sheet = sess.get("sheet")
        mode = sess.get("mode", "normal")
        if mode == "normal":
            if not sheet or sheet not in self.sheets:
                return
            self.current_qs = list(self.sheets[sheet])
            self.current_sheet = sheet
            self.mode = "normal"
        else:
            if not sheet or sheet not in self.sheets:
                return
            self.current_qs = self.db.wrong_list(sheet)
            self.current_sheet = sheet
            self.mode = "wrong"
        if not self.current_qs:
            if self.mode == "wrong":
                self.on_return()
            return
        qids = {q.qid for q in self.current_qs}
        saved_order = sess.get("order", [])
        saved_pos = min(max(sess.get("pos", 0), 0), max(0, len(saved_order) - 1))
        saved_current = saved_order[saved_pos] if saved_order else None
        order = [q for q in saved_order if q in qids]
        if mode == "wrong":
            # 错题集合可能在别处新增；恢复旧顺序后把新错题追加进来。
            order.extend(q.qid for q in self.current_qs if q.qid not in order)
        if not order:
            return
        self.order = order
        self.qmap = {q.qid: q for q in self.current_qs}
        self._round_order = sess.get("order_mode", "random")
        if self.mode == "normal":
            self._set_order_display(self._round_order)
        self.answers = {q: a for q, a in sess.get("answers", {}).items() if q in order}
        self._sync_scores()
        self.pos = (self.order.index(saved_current) if saved_current in self.order
                    else min(saved_pos, len(self.order) - 1))
        self._restored = True
        self._summary_mode = False
        self.placeholder.grid_remove()
        self.sheet_var.set(self.current_sheet if mode == "normal"
                           else f"❌ {self.current_sheet} · 错题重做")
        self._set_view("quiz")
        self._render()
        self._refresh_navigation()
        self._refresh_order_box()
        self._save_session()

    def _sync_scores(self):
        records = [self.answers[qid] for qid in set(self.order) if qid in self.answers]
        self.score_correct = sum(bool(rec.get("ok")) for rec in records)
        self.score_wrong = len(records) - self.score_correct

    def _accuracy_text(self):
        self._sync_scores()
        done = self.score_correct + self.score_wrong
        return f"{self.score_correct / done:.2%}" if done else "—"

    def _session_data(self):
        self._sync_scores()
        return {
            "sheet": self.current_sheet,
            "mode": self.mode,
            "order": self.order,
            "order_mode": self._round_order,
            "pos": self.pos,
            "answers": self.answers,
            "score": {"correct": self.score_correct, "wrong": self.score_wrong},
        }

    def _save_session(self):
        """保存当前及返回正常练习所需的进度。"""
        if not self.order:
            return
        sess = self._session_data()
        if self.mode == "normal" and self.current_sheet:
            self._sheet_sessions[self.current_sheet] = copy.deepcopy(sess)
        elif self.mode == "wrong" and self.current_sheet:
            self._wrong_sessions[self.current_sheet] = copy.deepcopy(sess)
        sess["normal_session"] = self._normal_session
        sess["sheet_sessions"] = self._sheet_sessions
        sess["wrong_sessions"] = self._wrong_sessions
        save_session(SESSION_FILE, sess)

    def _refresh_navigation(self):
        wrong = self.mode == "wrong"
        self.btn_restart.config(state="normal" if self.order else "disabled")
        self.sheet_combo.config(state="disabled" if wrong or not self.sheet_names else "readonly")
        self.btn_wrong.config(state="disabled" if wrong else "normal")
        self.btn_next.config(text="下一题")
        if wrong:
            self.btn_return.pack(side="left", padx=10)
        else:
            self.btn_return.pack_forget()

    def on_return(self):
        """退出错题模式，恢复进入前的题库、顺序和答题位置。"""
        self._cancel_auto_next()
        previous = self._normal_session
        self._summary_mode = False
        if previous and previous.get("sheet") in self.sheets:
            self._restore_session(previous)
        elif self.sheet_names:
            self.start_session(self.sheet_names[0])
        self._refresh_navigation()

    def on_restart(self):
        """重新开始当前练习，其他题库进度与长期错题本保留。"""
        if not self.order:
            return
        if not self.dialogs.askyesno("重新刷题", "确定重新开始当前练习？\n本轮答题记录将清空，其他题库进度和错题本保留。"):
            return
        self.start_session(self.current_sheet or "", wrong_only=self.mode == "wrong")

    def _on_close(self):
        self.root.after_cancel(self._ai_poll_id)
        if self._theme_watch_id is not None:
            self.root.after_cancel(self._theme_watch_id)
        self._save_session()
        self.db.close()
        self.root.destroy()

    def _fit_q_canvas(self, event=None):
        """左栏滚动区宽度变化：内容框跟随宽度，题干/解析按实际宽度换行"""
        if event is not None and event.widget is not self.q_canvas:
            return
        w = self.q_canvas.winfo_width()
        if w <= 20:
            return
        if getattr(self, "_last_q_w", 0) != w:
            self._last_q_w = w
            self.q_canvas.itemconfig(self.q_win, width=w)
            style = ttk.Style()
            # 按题目栏的真实可用宽度换行。不能设置较大的最小值，否则
            # 三栏布局或窄窗口下，题干会按比视口更宽的尺寸排版并被裁掉。
            wrap = max(80, w - 32)
            style.configure("Question.TLabel", wraplength=wrap)
            style.configure("Explain.TLabel", wraplength=wrap)
            # 直接设置控件，避免主题或显式字体导致样式中的 wraplength
            # 没有及时应用到已经创建的题干标签。
            self.question_label.configure(wraplength=wrap)
        self.root.after_idle(self._sync_q_scroll)

    def _bind_question_wheel(self, widget):
        def scroll(event):
            number = getattr(event, "num", None)
            if number in (4, 5):
                units = -3 if number == 4 else 3
            else:
                delta = event.delta
                if self.root.tk.call("tk", "windowingsystem") == "aqua":
                    units = -int(delta)
                else:
                    remainder = getattr(widget, "_wheel_delta", 0) + delta
                    ticks = int(remainder / 120)
                    widget._wheel_delta = remainder - ticks * 120
                    units = -ticks * 3
            target = self.q_canvas
            if isinstance(widget, tk.Text):
                top, bottom = widget.yview()
                if (units < 0 and top > 0) or (units > 0 and bottom < 1):
                    target = widget
            if units:
                target.yview_scroll(units, "units")
            return "break"
        for event in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            widget.bind(event, scroll)
        for child in widget.winfo_children():
            self._bind_question_wheel(child)

    def _sync_q_scroll(self):
        """内容短于视口时撑满（不留大片空白），内容超出时按内容高度可滚动。
        按子控件 reqheight 求和计算（canvas 窗口高度会钳制 winfo_height，不可用）"""
        try:
            cvh = self.q_canvas.winfo_height()
            total = self.q_content.winfo_reqheight()
        except tk.TclError:
            return
        h = max(total, cvh)
        cur = self.q_canvas.itemcget(self.q_win, "height")
        if cur in ("", None) or abs(int(cur) - h) > 1:
            self.q_canvas.itemconfig(self.q_win, height=h)
        w = self._last_q_w or self.q_canvas.winfo_width()
        cur_r = self.q_canvas.cget("scrollregion")
        if not cur_r or abs(float(cur_r.split()[3]) - h) > 1:
            self.q_canvas.config(scrollregion=(0, 0, w, h))

    # ---------------- 题目渲染 ----------------

    def _render(self):
        self.root.after_idle(self._install_shortcut_bindings)
        self._refresh_ai()
        self.root.after_idle(lambda: self._bind_question_wheel(self.q_content))
        if not self.order:
            return
        self.placeholder.grid_remove()
        if not self._q_shown:
            self.question_label.grid(row=1, column=0, sticky="ew", padx=4, pady=(12, 20))
            self.options_frame.grid(row=2, column=0, sticky="ew", padx=4)
            self.btn_submit.grid(row=3, column=0, sticky="w", padx=8, pady=(4, 0))
            self.result_label.grid(row=4, column=0, sticky="ew", padx=4, pady=4)
            self.explain_label.grid(row=1, column=0, sticky="ew", padx=0, pady=(0, 8))
            self.nav_frame.grid(row=0, column=0, sticky="w", padx=0, pady=(10, 12))
            self._q_shown = True

        q = self.qmap.get(self.order[self.pos]) or self.current_qs[0]
        self.btn_favorite.config(text="★ 已收藏 · 取消" if self.db.is_favorite(q.qid) else "☆ 收藏题目")
        # 题干
        tag = f"【{q.type}】  {q.difficulty}"
        if self.mode == "wrong":
            tag += "  （错题重做）"
        self.question_label.config(text=f"{q.stem}")
        self.root.after_idle(self._fit_q_canvas)
        self.question_meta.config(text=f"第 {self.pos + 1} / {len(self.order)} 题    {tag}")
        self.question_meta.grid(row=0, column=0, sticky="w", padx=4, pady=(4, 0))

        # 清空旧选项
        for w in self.options_frame.winfo_children():
            w.destroy()
        self._multi_vals = []

        rec = self.answers.get(q.qid)
        answered = rec is not None
        prev = list(rec["sel"]) if answered else []

        # 单选 / 判断 用单选按钮（判断题只允许选一个）；多选 用复选框
        var = tk.StringVar()
        self._cur_var = var
        self._multi_vals = []
        for opt in q.options:
            letter = opt[0]
            row = tk.Frame(self.options_frame, bg=BG_CARD, highlightthickness=1,
                           highlightbackground=BORDER)
            row.pack(anchor="w", fill="x", pady=6, padx=2)
            if q.type == "多选":
                v = tk.BooleanVar(value=(letter in prev))
                self._multi_vals.append((letter, v))
                cb = tk.Checkbutton(
                    row, text=opt, variable=v, font=_font(11),
                    bg=BG_CARD, fg=FG_TEXT, activebackground=BG_SUBTLE,
                    activeforeground=FG_TEXT, anchor="w",
                    highlightthickness=0, selectcolor=BG_CARD,
                    state="disabled" if answered else "normal",
                )
            else:
                cb = tk.Radiobutton(
                    row, text=opt, value=letter, variable=var, font=_font(11),
                    indicatoron=False, relief="flat", offrelief="flat", borderwidth=0,
                    command=lambda letter=letter, q=q: self._click_instant(q, letter),
                    bg=BG_CARD, fg=FG_TEXT, activebackground=BG_SUBTLE,
                    activeforeground=FG_TEXT, anchor="w",
                    highlightthickness=0, selectcolor=BG_CARD,
                    state="disabled" if answered else "normal",
                )
                if letter in prev:
                    var.set(letter)
            cb.pack(side="left", fill="x", expand=True, padx=16, pady=13)
            if not answered and q.type in ("单选", "判断"):
                # 点选即判分（多选保留提交按钮）
                row.bind("<Button-1>", lambda e, letter=letter, q=q: self._click_instant(q, letter))
                # 悬停高亮整行
                row.bind("<Enter>", lambda e, r=row: r.config(highlightbackground=ACCENT))
                row.bind("<Leave>", lambda e, r=row: r.config(highlightbackground=BORDER))
            self._paint_option(row, cb, letter, q, prev, answered)

        # 结果条（内联显示对错；答错只提示选项字母，不带选项内容）
        if answered:
            self.result_label.grid()
            ok = rec["ok"]
            correct_letters = "、".join(sorted(q.answer))
            if ok:
                self.result_label.config(text="✅ 回答正确", style="Ok.TLabel")
            else:
                sel_text = "、".join(prev) if prev else "未作答"
                self.result_label.config(text=f"❌ 回答错误（你选了 {sel_text}）　正确答案：{correct_letters}",
                                         style="Bad.TLabel")
        else:
            self.result_label.config(text="", style="Sub.TLabel")
            self.result_label.grid_remove()
        # 解析
        if q.explanation:
            self.explain_label.config(text=f"解析：{q.explanation}")
        else:
            self.explain_label.config(text="")

        # 按钮状态：多选需要提交按钮（选项正下方）；单选/判断点选即判，隐藏提交按钮
        if q.type == "多选":
            self.btn_submit.grid()
            self.btn_submit.config(state="disabled" if answered else "normal")
        else:
            self.btn_submit.grid_remove()
        self.btn_prev.config(state="normal" if self.pos > 0 else "disabled")
        self.btn_next.config(state="normal" if self.pos < len(self.order) - 1 else "disabled")

        # 进度（按已答题数）
        accuracy = self._accuracy_text()
        done = self.score_correct + self.score_wrong
        self.progress_var.set(int(done / max(1, len(self.order)) * 100))
        self.progress_text.config(text=f"已答 {done} / {len(self.order)}　准确率 {accuracy}")
        self.status_label.config(text=f"第 {self.pos + 1} 题  ·  共 {len(self.order)} 题")
        # 矩阵平时只重着色（切题），题目列表变了才重建
        if len(self.list_cells) != len(self.order):
            self._rebuild_matrix()
        else:
            self._paint_matrix()
        # 内容高度随题目变化 → 重新同步滚动区（短题撑满、长题可滚动）
        self.root.after_idle(self._sync_q_scroll)

    def _q_at(self, pos: int) -> Question:
        return self.qmap.get(self.order[pos]) or self.current_qs[0]

    def _collect_selection(self) -> list[str]:
        if self._multi_vals:
            return [letter for letter, var in self._multi_vals if var.get()]
        v = self._cur_var.get()
        return [v] if v else []

    # ---------------- 题目矩阵 ----------------

    def _on_matrix_wheel(self, event):
        """只滚动鼠标所在的答题卡，兼容高精度滚轮和 Linux。"""
        if getattr(event, "num", None) in (4, 5):
            steps = -1 if event.num == 4 else 1
        elif self.root.tk.call("tk", "windowingsystem") == "aqua":
            steps = -int(event.delta)
        else:
            self._matrix_wheel_delta += event.delta
            ticks = int(self._matrix_wheel_delta / 120)
            self._matrix_wheel_delta -= ticks * 120
            steps = -ticks * 3
        if steps and self.list_canvas.yview() != (0.0, 1.0):
            self.list_canvas.yview_scroll(steps, "units")
        return "break"

    def _fit_list_matrix(self, event=None):
        """画布宽度变化时调整矩阵列数（仅列数变化才重建，防循环）"""
        if event is not None and event.widget is not self.list_canvas:
            return
        w = self.list_canvas.winfo_width()
        if w <= 20 or getattr(self, "_last_fit_w", 0) == w:
            return
        self._last_fit_w = w
        new_cols = max(4, min(12, w // (self._cell_w + 2)))
        if self.order and new_cols != self._list_cols:
            self._rebuild_matrix(cols=new_cols)

    def _matrix_color(self, qid: str, i: int) -> str:
        """按状态取色：当前题 > 答对 > 答错 > 未答"""
        if i == self.pos:
            return C_CURRENT
        rec = self.answers.get(qid)
        if rec is None:
            return C_UNANSWERED
        return C_OK if rec["ok"] else C_BAD

    def _rebuild_matrix(self, cols=None):
        """(重新)构建矩阵：单 Canvas，每题一个矩形+序号，只在题目列表/列数变化时调用"""
        if not self.order:
            return
        w = self.list_canvas.winfo_width()
        if cols is None:
            cols = max(4, min(12, w // (self._cell_w + 2))) if w > 20 else 8
        self._list_cols = cols
        self.list_canvas.delete("all")
        self.list_cells = []
        self._cell_colors = []
        cw, ch = self._cell_w, self._cell_h
        for i in range(len(self.order)):
            x0, y0 = (i % cols) * cw, (i // cols) * ch
            r = self.list_canvas.create_rectangle(x0 + 4, y0 + 4, x0 + cw - 4, y0 + ch - 4,
                                                   fill=C_UNANSWERED, outline=BORDER_SOFT)
            t = self.list_canvas.create_text(x0 + cw // 2, y0 + ch // 2, text=str(i + 1),
                                              font=_font(10), fill=C_TEXT_ON_LIGHT)
            self.list_cells.append((r, t))
            self._cell_colors.append(C_UNANSWERED)
        self._list_rows = (len(self.order) + cols - 1) // cols
        self.list_canvas.config(scrollregion=(0, 0, cols * cw, self._list_rows * ch))
        self._paint_matrix()

    def _render_list(self):
        """兼容旧调用：题目变化时重建矩阵，否则只重着色"""
        if len(self.list_cells) != len(self.order or []):
            self._rebuild_matrix()
        else:
            self._paint_matrix()

    def _paint_matrix(self):
        """按当前状态重着色（仅更新颜色发生变化的格子，答一题只改 2~3 格）"""
        dark_cells = {C_CURRENT, C_BAD, C_OK}  # 深色格子 → 白字
        for i, (r, t) in enumerate(self.list_cells):
            color = self._matrix_color(self.order[i], i)
            if color == self._cell_colors[i]:
                continue
            self._cell_colors[i] = color
            self.list_canvas.itemconfig(r, fill=color)
            self.list_canvas.itemconfig(t,
                                        fill=C_TEXT_ON_CELL if color in dark_cells
                                        else C_TEXT_ON_LIGHT)

    def _on_matrix_click(self, event):
        """点击矩阵 → 按坐标命中题号并跳转"""
        if not self.order or not self.list_cells:
            return
        cw, ch, cols = self._cell_w, self._cell_h, self._list_cols
        total_h = self._list_rows * ch
        vy = self.list_canvas.yview()[0] * total_h if total_h > self.list_canvas.winfo_height() else 0
        idx = int((event.y + vy) // ch) * cols + int(event.x // cw)
        if 0 <= idx < len(self.order):
            self._jump_to(self.order[idx])

    def _jump_to(self, qid: str):
        """点击矩阵格 → 跳转到该题"""
        if qid not in self.order:
            return
        self.pos = self.order.index(qid)
        self._cancel_auto_next()
        self._render()
        self._save_session()

    def _paint_option(self, row, cb, letter: str, q: Question,
                      prev: list[str], answered: bool):
        """已答题：整行着色——正确选项淡绿，选错的淡红，其余淡灰"""
        if not answered:
            return
        try:
            if letter in q.answer:
                row.config(bg=C_OK_SOFT, highlightbackground=C_OK)
                cb.config(fg=C_OK, disabledforeground=C_OK, bg=C_OK_SOFT)
            elif letter in prev:
                row.config(bg=C_BAD_SOFT, highlightbackground=C_BAD)
                cb.config(fg=C_BAD, disabledforeground=C_BAD, bg=C_BAD_SOFT)
            else:
                row.config(bg=BG_CARD, highlightbackground=BORDER)
                cb.config(fg=FG_FAINT, disabledforeground=FG_MUTED)
        except tk.TclError:
            pass

    # ---------------- 上一题 / 下一题 ----------------

    def _cancel_auto_next(self):
        """取消已排定的自动跳转（用户手动切题时）"""
        if self._pending is not None:
            try:
                self.root.after_cancel(self._pending)
            except (tk.TclError, ValueError):
                pass
            self._pending = None

    def on_prev(self):
        self._cancel_auto_next()
        if self.pos <= 0:
            return
        self.pos -= 1
        self._render()
        self._save_session()

    def on_next(self):
        self._cancel_auto_next()
        if self.pos >= len(self.order) - 1:
            if getattr(self, "_summary_mode", False):
                # 总结页：重新开始本轮
                self._summary_mode = False
                self.btn_next.config(text="下一题 →")
                if self.mode == "wrong":
                    self.start_session(self.current_sheet, wrong_only=True)
                else:
                    self.start_session(self.current_sheet)
                return
            # 全部答完：内联显示总结
            self._show_summary()
            return
        self.pos += 1
        self._render()
        self._save_session()

    def _show_summary(self):
        accuracy = self._accuracy_text()
        total = len(self.order)
        correct = self.score_correct
        pct = correct / total * 100 if total else 0
        self._summary_mode = True
        self.question_label.config(
            text=f"本轮完成！\n\n共 {total} 题，正确 {correct} 题，错误 {self.score_wrong} 题，准确率 {accuracy}")
        for w in self.options_frame.winfo_children():
            w.destroy()
        self.options_frame.pack_forget()
        self.result_label.config(text="🎉 全部完成", style="Ok.TLabel")
        self.explain_label.config(text="点「下一题」或切换试卷继续刷题；错题可点「重做错题」。")
        self.btn_next.config(text="重新开始本轮")
        self._render_list()
        self._save_session()

    # ---------------- 判分（内联显示） ----------------

    def _grade(self, q: Question, sel: list[str]) -> bool:
        return set(sel) == set(q.answer)

    def _click_instant(self, q: Question, letter: str):
        """单选/判断：点选项即判分"""
        if q.qid in self.answers:
            return
        self._apply_answer(q, [letter])

    def on_submit(self):
        """多选：点「提交本题」判分"""
        if not self.order:
            return
        q = self._q_at(self.pos)
        if q.qid in self.answers:
            return
        self._apply_answer(q, self._collect_selection())

    def _apply_answer(self, q: Question, sel: list[str]):
        """记录答案、判分、更新错题本、内联显示，并安排自动跳转"""
        self._cancel_auto_next()
        ok = self._grade(q, sel)
        self.answers[q.qid] = {"sel": sel, "ok": ok}
        if ok:
            self.score_correct += 1
            # 答对：从长期错题本移除
            self.db.wrong_remove(q.qid)
        else:
            self.score_wrong += 1
            self.db.wrong_add(q.qid)
        self._render()
        self._save_session()
        if ok and self.pos < len(self.order) - 1:
            # 答对：短暂停顿后自动进入下一题
            self._pending = self.root.after(self._auto_next_delay, self._auto_next)

    def _auto_next(self):
        """答对后的自动跳转（若期间用户手动切题则已取消，不再触发）"""
        self._pending = None
        if self.pos < len(self.order) - 1:
            self.pos += 1
            self._render()
            self._save_session()

    # ---------------- 错题重做 ----------------

    def on_wrong(self):
        sheet = self.current_sheet
        wrong = self.db.wrong_list(sheet) if sheet else []
        if not wrong:
            self.dialogs.showinfo("提示", f"「{sheet or '当前题库'}」暂无错题，继续加油！")
            return
        if not self.dialogs.askyesno("重做错题",
                                   f"「{sheet}」错题本中有 {len(wrong)} 道题，是否重做？"):
            return
        saved = self._wrong_sessions.get(sheet)
        if saved:
            # 先记录返回普通练习所需的现场，再恢复该题库错题进度。
            if self.mode == "normal" and self.order:
                self._normal_session = self._session_data()
            self._restore_session(copy.deepcopy(saved))
        else:
            self.start_session(sheet, wrong_only=True)
        self.status_label.config(text=f"{sheet} · 错题重做")

    # ---------------- 试卷切换 ----------------

    def _on_sheet_change(self, event=None):
        sheet = self.sheet_var.get()
        if not sheet or sheet == "❌ 错题重做" or sheet not in self.sheets:
            return
        if sheet == self.current_sheet and self.mode == "normal":
            return
        self._cancel_auto_next()
        self._save_session()
        saved = self._sheet_sessions.get(sheet)
        valid_ids = {q.qid for q in self.sheets[sheet]}
        if saved and any(qid in valid_ids for qid in saved.get("order", [])):
            self._restore_session(copy.deepcopy(saved))
        else:
            self.start_session(sheet)

    # ---------------- 工具 ----------------

    def _toast(self, msg: str):
        self.status_label.config(text=msg)


def main():
    root = tk.Tk()
    # 高 DPI 支持
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    QuizApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
