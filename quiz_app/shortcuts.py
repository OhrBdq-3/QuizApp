"""Configurable quiz shortcuts, independent of model settings."""
import json
import os
import time
import tkinter as tk
from tkinter import ttk

DEFAULT_KEYS = {'A': 'a', 'B': 'b', 'C': 'c', 'D': 'd', 'submit': 'Return'}

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
        for child in widget.winfo_children():
            self._install_shortcut_bindings(child)

    def _answer_shortcut(self, event):
        if getattr(self, '_current_view', '') != 'quiz' or not self.order or self._summary_mode:
            return
        if event.widget.winfo_toplevel() != self.root or self.root.grab_current() is not None:
            return
        if event.widget.winfo_class() in ('Entry', 'TEntry', 'TCombobox', 'Text', 'Spinbox', 'TSpinbox'):
            return
        if event.state & (0x4 | 0x8 | 0x20000):
            return
        key = event.keysym.lower()
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
        ttk.Label(body, text='支持字母、数字、Enter、Space。\n单选直接作答；多选切换勾选，再按提交键。\n编辑输入框和阅读解析时不触发答题。').grid(row=5, column=0, columnspan=2, pady=16)
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
        ttk.Button(body, text='恢复默认', command=lambda: [variables[a].set(k) for a, k in DEFAULT_KEYS.items()]).grid(row=6, column=0)
        ttk.Button(body, text='保存快捷键', command=save).grid(row=6, column=1)
