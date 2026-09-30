# -*- coding: utf-8 -*-
"""智能筛题（背题规则）
====================
把 build_memorization_bank.py 的"需要背的题"筛选规则移植到 App 内。

规则（与脚本/notebook 口径一致，关键词沿用子串匹配，包括"仅"命中"不仅"）：
  - 判断：题干命中绝对化关键词 且 标准答案为「错误」→ 剔除（背）；
          字母答案按选项文本解析（A=正确 / B=错误）。
  - 多选：至少有一个错误选项，且每个错误选项都命中关键词 → 剔除；
          全选题（无错误选项）保留。正确选项含关键词不影响判断。
  - 单选：默认仅剔除「正确答案唯一最长」的题（--single-rule keep 可保留全部）。

对异常数据（题干为空、答案无法解析、选项不足等）一律保留，不剔除。
"""

from __future__ import annotations

import re
import unicodedata

ABSOLUTE_WORDS = [
    "仅", "无需", "一定", "始终", "绝对", "不能", "全部", "完全", "无限制",
    "任何", "不做任何", "任何情况下", "必然", "必须", "绝对不能",
    "没有任何", "完全不", "毫无", "随意", "永远",
]
FALSE_LABELS = {"错", "错误", "不正确", "否", "false", "f", "×", "✗", "✘"}
TRUE_LABELS = {"对", "正确", "是", "true", "t", "√", "✓", "✔"}


def _text(v) -> str:
    return "" if v is None else str(v).strip()


def _normalize(v) -> str:
    return unicodedata.normalize("NFKC", _text(v))


def _keyword_hits(value) -> list[str]:
    return [w for w in ABSOLUTE_WORDS if w in _text(value)]


def _option_body(opt: str) -> str:
    """'A. xxx' -> 'xxx'"""
    return re.sub(r"^[A-Za-z]\s*[.、．:：\)）]\s*", "", opt or "").strip()


def _option_letter(opt: str) -> str:
    return (opt or "")[:1].upper()


def should_drop(q) -> tuple[bool, str]:
    """判断单道题是否需要背（剔除）。

    返回 (是否剔除, 原因)。异常数据一律保留（剔除=False）。
    """
    stem = _text(q.stem)
    if not stem:
        return False, "题干为空"

    options = [o for o in q.options if _text(o)]
    answers = [a.upper() for a in q.answer if a]
    letters = {_option_letter(o) for o in options}

    qtype = (q.type or "").replace(" ", "")
    if qtype in ("判断", "判断题"):
        # 判断题：字母答案按选项文本解析；中文答案直接判
        value = _normalize(q.answer[0]) if answers else ""
        if value.upper() in letters:  # 写的是字母，映射到选项文字
            value = _normalize(_option_body(
                next((o for o in options if _option_letter(o) == value.upper()), "")))
        if value.strip().lower() not in (FALSE_LABELS | TRUE_LABELS):
            return False, "无法解析判断题答案"
        value = value.strip().lower()
        hits = _keyword_hits(stem)
        if value in FALSE_LABELS and hits:
            return True, "判断答案为错误；题干命中：" + "、".join(hits)
        return False, ""

    if qtype not in ("单选", "单选题", "多选", "多选题"):
        return False, "未支持的题型，已保留"
    if not answers or not set(answers) <= letters:
        return False, "答案为空、格式异常或指向空选项"
    if len(options) < 2:
        return False, "有效选项少于两个"

    if qtype in ("多选", "多选题"):
        if len(set(answers)) < 2:
            return False, "多选题正确答案少于两个"
        incorrect = [o for o in options if _option_letter(o) not in answers]
        if not incorrect:
            return False, "全选题保留"
        hits = {_option_letter(o): _keyword_hits(_option_body(o)) for o in incorrect}
        if all(hits.values()):
            reason = "所有错误选项均命中：" + "；".join(
                f"{c}（{'、'.join(w)}）" for c, w in hits.items())
            return True, reason
        return False, ""

    # 单选
    if len(set(answers)) != 1:
        return False, "单选题答案不是单个选项"
    lengths = {_option_letter(o): len(_option_body(o)) for o in options}
    longest = [c for c, ln in lengths.items() if ln == max(lengths.values())]
    # 唯一最长，且正确答案就是它 → 剔除
    if len(longest) == 1 and longest[0] in answers:
        return True, "正确答案唯一最长"
    return False, ""


def filter_questions(qs) -> tuple[list, list]:
    """对一组 Question 应用背题规则。

    返回 (需要背的题[原顺序], 被剔除的[(q, 原因), ...])。
    """
    kept: list = []
    removed: list[tuple] = []
    for q in qs:
        drop, reason = should_drop(q)
        if drop:
            removed.append((q, reason))
        else:
            kept.append(q)
    return kept, removed
