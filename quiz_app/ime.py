# -*- coding: utf-8 -*-
"""Windows 输入法（IME）处理。

问题背景
--------
中文输入法处于中文状态时，字母键会被 IME 截走：应用收到的 KeyPress 事件
是 ``keysym='??'``、``keycode=229``（VK_PROCESSKEY），而不是 'a'/'s'/'d'/'f'。
结果是答题快捷键「按了没反应」——代码逻辑没问题，是按键根本没送到。

解决办法
--------
给窗口摘掉 IME 关联（``ImmAssociateContext(hwnd, NULL)``），字母键就会以
原始虚拟键码送达，Tk 能正常拿到 keysym。该设置只作用于本应用的窗口，
不影响用户在其它程序里的输入法状态。

非 Windows 平台（macOS / Linux 没有这套机制）全部函数直接返回 False，可安全调用。
"""
from __future__ import annotations

import sys

SUPPORTED = sys.platform == "win32"

# VK_PROCESSKEY：按键被输入法处理时 Tk 报告的这个 keycode
IME_KEYCODE = 229

_imm32 = None


def _imm():
    """惰性加载 imm32.dll（非 Windows 或加载失败返回 None）"""
    global _imm32
    if not SUPPORTED:
        return None
    if _imm32 is None:
        try:
            import ctypes
            from ctypes import wintypes

            dll = ctypes.WinDLL("imm32")
            HIMC = ctypes.c_void_p
            dll.ImmAssociateContext.restype = HIMC
            dll.ImmAssociateContext.argtypes = [wintypes.HWND, HIMC]
            dll.ImmGetContext.restype = HIMC
            dll.ImmGetContext.argtypes = [wintypes.HWND]
            _imm32 = (dll, ctypes, wintypes, HIMC)
        except Exception:
            _imm32 = False      # 标记：试过但不可用
    return _imm32 or None


def supported() -> bool:
    return SUPPORTED and _imm() is not None


def disable(hwnd: int) -> bool:
    """摘掉某个窗口句柄的 IME 关联。返回是否成功。"""
    loaded = _imm()
    if not loaded or not hwnd:
        return False
    dll, ctypes, wintypes, HIMC = loaded
    try:
        dll.ImmAssociateContext(wintypes.HWND(int(hwnd)), HIMC(0))
        return True
    except Exception:
        return False


def is_ime_key(event) -> bool:
    """判断这个按键事件是不是被输入法吞掉了。"""
    try:
        return getattr(event, "keycode", 0) == IME_KEYCODE or event.keysym == "??"
    except Exception:
        return False
