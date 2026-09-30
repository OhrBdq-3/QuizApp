"""Configurable quiz shortcuts, independent of model settings."""
import json
import os
import sys
import time
import tkinter as tk
from tkinter import ttk

from . import ime

DEFAULT_KEYS = {'A': 'a', 'B': 'b', 'C': 'c', 'D': 'd', 'submit': 'Return'}

# 翻页键：固定绑方向键，不参与自定义设置——方向键翻页是通用直觉，
# 让用户改反而容易改坏。想改的话直接改这里即可（小写 keysym）。
NAV_KEYS = {'prev': ('left', 'up'), 'next': ('right', 'down')}

# 给界面/文档用的说明文案
NAV_HINT = "← ↑ 上一题    → ↓ 下一题"

def validate_keys(values):
    result = {}
    for action in DEFAULT_KEYS:
        key = values[action].strip()
        if key.lower() in ('enter', 'return'):
            key = 'Return'
        elif key.lower() == 'space':
            key = 'space'
        elif len(key) == 1 and key.isascii() and key.isalnum():
            key = key.lower()
        else:
            raise ValueError('快捷键支持单个英文字母、数字、Enter 或 Space。')
        result[action] = key
    if len(set(result.values())) != len(result):
        raise ValueError('快捷键不能重复。')
    return result

class ShortcutMixin:
    def _init_shortcuts(self, path):
        self._shortcut_path = path
        self._shortcut_keys = dict(DEFAULT_KEYS)
        self._held_keys = set()
        try:
            with open(path, encoding='utf-8') as f:
                self._shortcut_keys = validate_keys(json.load(f))
        except (OSError, ValueError, KeyError, TypeError):
            pass
        self._shortcut_tag = 'QuizShortcuts' + str(id(self))
        self.root.bind_class(self._shortcut_tag, '<KeyPress>', self._answer_shortcut)
        self.root.bind_class(self._shortcut_tag, '<KeyRelease>', lambda e: self._held_keys.discard(e.keysym.lower()))
        self.root.bind('<FocusOut>', lambda e: self._held_keys.clear(), add='+')

    def _install_shortcut_bindings(self, widget=None):
        widget = self.root if widget is None else widget
        if widget.winfo_toplevel() != self.root:
            return
        tags = widget.bindtags()
        if self._shortcut_tag not in tags:
            widget.bindtags((self._shortcut_tag,) + tags)
        self._strip_ime(widget)
        for child in widget.winfo_children():
            self._install_shortcut_bindings(child)

    def _strip_ime(self, widget):
        """给非输入控件摘掉输入法关联。

        中文输入法开启时字母键会被 IME 吞掉（keysym 变成 '??'），
        快捷键就完全没反应。摘掉关联后按键以原始键码送达，快捷键才生效。
        真正的输入框保留 IME，不影响需要打字的地方。
        """
        if not ime.supported():
            return
        stripped = getattr(self, "_ime_stripped", None)
        if stripped is None:
            stripped = self._ime_stripped = set()
        try:
            hwnd = widget.winfo_id()
        except Exception:
            return
        if not hwnd or hwnd in stripped:
            return
        stripped.add(hwnd)
        if self._is_text_input(widget):
            return
        ime.disable(hwnd)

    # 会吃掉按键的“真输入框”：只有真正能敲进文字的控件才拦掉快捷键。
    # 只读下拉框（state=readonly）和只读文本框（state=disabled，例如 AI 解析面板）
    # 不在此列——点过它们之后焦点会留驻，若一律拦截，快捷键就会看起来“失灵”。
    _ALWAYS_INPUT = ('Entry', 'TEntry', 'Spinbox', 'TSpinbox')

    # 方向键在这些控件里有原生用途：Text 用来滚动长文、下拉框用来改选中项。
    # 抢过来的话解析面板就没法用键盘翻页了，所以这些控件里翻页键一律放行。
    _NAV_NATIVE = ('Text', 'TCombobox', 'Entry', 'TEntry', 'Spinbox', 'TSpinbox')

    def _is_text_input(self, widget) -> bool:
        try:
            cls = widget.winfo_class()
        except Exception:
            return False
        if cls in self._ALWAYS_INPUT:
            return True
        if cls in ('Text', 'TCombobox'):
            try:
                return str(widget.cget('state')).lower() == 'normal'
            except Exception:
                return False
        return False

    @staticmethod
    def _modifier_held(event) -> bool:
        """是否真的按住了 Ctrl / Alt 组合键。

        Tk 报的 event.state 并不完全可信：某些环境（远程桌面、注入式按键、
        部分输入法状态）下普通字母键也会带上 Alt 位（0x8），一律照挡就会把
        所有快捷键误杀。所以只在 state 提示有修饰键时，再用系统 API 复核一次。
        """
        state = getattr(event, "state", 0) or 0
        if not state & (0x4 | 0x8):        # 0x4=Control，0x8=Alt
            return False
        if sys.platform != "win32":
            return True                    # 非 Windows 只能相信 event.state
        try:
            import ctypes
            user32 = ctypes.windll.user32
            VK_CONTROL, VK_MENU = 0x11, 0x12
            return bool(user32.GetAsyncKeyState(VK_CONTROL) & 0x8000) or \
                   bool(user32.GetAsyncKeyState(VK_MENU) & 0x8000)
        except Exception:
            return True

    def _hint_ime(self):
        """按键被输入法吞掉时给个明确提示，避免“快捷键没反应”无从下手。"""
        msg = "输入法处于中文状态，答题快捷键不生效——请切换到英文输入（Ctrl+空格 / Shift）"
        toast = getattr(self, "_toast", None)
        if callable(toast):
            try:
                toast(msg)
            except Exception:
                pass

    def _nav_shortcut(self, event, key: str) -> bool:
        """方向键翻页：← ↑ 上一题，→ ↓ 下一题。

        放在答题键之前处理，这样答完题也能翻页回看（答题键在已作答时会跳过）。
        返回 True 表示这个按键已被翻页消费掉，调用方应当 'break'。
        """
        action = next((a for a, keys in NAV_KEYS.items() if key in keys), None)
        if action is None:
            return False
        try:
            if event.widget.winfo_class() in self._NAV_NATIVE:
                return False        # 让 Text 滚动、下拉框改值等原生行为照旧
        except Exception:
            pass
        if key in self._held_keys:
            return True             # 长按只翻一题，不连跳
        self._held_keys.add(key)
        if action == 'prev':
            # 总结页没有「上一题」，回到题目列表会留下半总结状态，直接忽略
            if getattr(self, '_summary_mode', False) or self.pos <= 0:
                return True
            self.on_prev()
        else:
            self.on_next()          # 最后一题时走既有逻辑：显示总结 / 重新开始
        return True

    def _answer_shortcut(self, event):
        if getattr(self, '_current_view', '') != 'quiz' or not self.order:
            return
        if event.widget.winfo_toplevel() != self.root or self.root.grab_current() is not None:
            return
        if self._is_text_input(event.widget):
            return
        if self._modifier_held(event):
            return
        if ime.is_ime_key(event):
            # 输入法仍开着（理论上已被 _strip_ime 摘掉，这里兜底提示）
            self._hint_ime()
            return
        key = event.keysym.lower()
        if self._nav_shortcut(event, key):
            return 'break'
        if getattr(self, '_summary_mode', False):
            return
        action = next((a for a, k in self._shortcut_keys.items() if k.lower() == key), None)
        if action is None:
            return
        if key in self._held_keys:
            return 'break'
        self._held_keys.add(key)
        q = self._q_at(self.pos)
        if q.qid in self.answers:
            return 'break'
        if action == 'submit':
            if q.type == '多选':
                self.on_submit()
        elif action in [option[0] for option in q.options]:
            if q.type == '多选':
                for letter, variable in self._multi_vals:
                    if letter == action:
                        variable.set(not variable.get())
            else:
                self._click_instant(q, action)
        return 'break'

    def open_shortcut_settings(self):
        win = tk.Toplevel(self.root)
        win.title('设置 · 答题快捷键')
        win.transient(self.root)
        win.resizable(False, False)
        win.configure(bg='#ffffff')
        body = ttk.Frame(win, padding=24)
        body.pack(fill='both', expand=True)
        variables = {}
        for row, action in enumerate(DEFAULT_KEYS):
            ttk.Label(body, text='多选提交' if action == 'submit' else f'选项 {action}').grid(row=row, column=0, padx=(0, 24), pady=8, sticky='w')
            variables[action] = tk.StringVar(value=self._shortcut_keys[action])
            ttk.Entry(body, textvariable=variables[action], width=18).grid(row=row, column=1)
        ttk.Label(body, text='支持字母、数字、Enter、Space。\n单选直接作答；多选切换勾选，再按提交键。\n编辑输入框和阅读解析时不触发答题。',
                  justify='left').grid(row=5, column=0, columnspan=2, pady=(16, 4), sticky='w')
        ttk.Label(body, text=f'翻页固定为方向键（不可改）：{NAV_HINT}',
                  justify='left').grid(row=6, column=0, columnspan=2, pady=(0, 16), sticky='w')
        def save():
            try:
                keys = validate_keys({a: v.get() for a, v in variables.items()})
                with open(self._shortcut_path + '.tmp', 'w', encoding='utf-8') as f:
                    json.dump(keys, f, ensure_ascii=False, indent=2)
                os.replace(self._shortcut_path + '.tmp', self._shortcut_path)
            except (ValueError, OSError) as exc:
                self.dialogs.showerror('无法保存', str(exc), parent=win)
                return
            self._shortcut_keys = keys
            self._held_keys.clear()
            win.destroy()
        ttk.Button(body, text='恢复默认', command=lambda: [variables[a].set(k) for a, k in DEFAULT_KEYS.items()]).grid(row=7, column=0)
        ttk.Button(body, text='保存快捷键', command=save).grid(row=7, column=1)
