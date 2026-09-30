# -*- coding: utf-8 -*-
"""Windows 标题栏染色（DWM）。

Tk 只能画客户区；带全屏/最小化/关闭按钮的那条标题栏归 Windows 的 DWM 管，
Tk 没有任何接口能改它的颜色。这里通过 DwmSetWindowAttribute 直接上色：

    DWMWA_USE_IMMERSIVE_DARK_MODE = 20   Win10 1903+（旧构建用 19）
    DWMWA_CAPTION_COLOR           = 35   Win11 22000+
    DWMWA_BORDER_COLOR            = 34
    DWMWA_TEXT_COLOR              = 36

配色只区分浅色/深色（THEME_COLORS），不跟主题色走；Win10 退化处理——
只设暗色标志，标题栏深浅跟随主题明暗。非 Windows 或 API 不可用时静默跳过。

COLORREF 是 0x00BBGGRR（注意不是 RGB），hex 需要翻转后两段。
"""
import ctypes

# ---------- 常量 ----------
_DWMWA_USE_IMMERSIVE_DARK_MODE = 20
_DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19
_DWMWA_BORDER_COLOR = 34
_DWMWA_CAPTION_COLOR = 35
_DWMWA_TEXT_COLOR = 36

_S_OK = 0

# 标题栏只区分浅色/深色，与应用背景同色系（不跟主题色走）。
#   light: 底色 = 应用底 BG_APP #f5f0e8，文字 = 主文字 FG_TEXT #29251f
#   dark:  底色 = 应用底 BG_APP #14120e，文字 = 主文字 FG_TEXT #ece5d8
# 标题栏底色刻意用"页面底"而不是"卡片色"：整窗从系统外框到客户区过渡更自然。
THEME_COLORS = {
    "light": {"caption": "#f5f0e8", "text": "#29251f"},
    "dark": {"caption": "#14120e", "text": "#ece5d8"},
}


def colors_for(resolved_theme):
    """按已解析的明暗主题（'light' / 'dark'）取标题栏配色。"""
    return THEME_COLORS.get(resolved_theme, THEME_COLORS["light"])


def _load():
    """懒加载 dwmapi/user32；任何一步失败都返回 None（整体降级为 no-op）。"""
    try:
        return ctypes.WinDLL("dwmapi"), ctypes.windll.user32
    except (AttributeError, OSError):      # 非 Windows 或缺少 dll
        return None, None


_dwm, _user32 = _load()


def _colorref(hex_color):
    """'#rrggbb' -> COLORREF 0x00BBGGRR。"""
    h = str(hex_color).lstrip("#")
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return r | (g << 8) | (b << 16)


def _luminance(hex_color):
    h = str(hex_color).lstrip("#")
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


def hwnd_of(widget):
    """Tk 控件 -> 真正的顶层窗口 HWND；拿不到时返回 0。

    winfo_id() 给的是 TkChild（客户区子窗口），带标题栏的 TkTopLevel 是它的
    父窗口（已实测：class 名分别为 TkChild / TkTopLevel）。

    坑：Tk 的顶层窗口是**延迟创建**的。刚 tk.Tk() 出来时尚未映射，父链上
    只有 TkChild、没有 TkTopLevel——此时绝不能把 TkChild 当成顶层返回，否则
    属性会写到客户区子窗口上（看着设成功了），随后 Tk 真建窗又把它冲回默认
    值。这种"以为设上了其实没生效"比直接失败更难查，所以这里明确返回 0，
    由调用方在窗口映射后重试（见 QuizApp._on_root_mapped）。
    """
    if _user32 is None:
        return 0
    try:
        wid = widget.winfo_id()
    except (AttributeError, ctypes.ArgumentError):
        return 0
    hwnd = wid
    for _ in range(4):
        cls = _class_name(hwnd)
        if cls == "TkTopLevel":
            return hwnd
        parent = _user32.GetParent(ctypes.c_void_p(hwnd))
        if not parent or parent == hwnd:
            break
        hwnd = parent
    # 没找到 TkTopLevel：窗口还没真正创建，返回 0 让调用方稍后重试
    return 0


def _class_name(hwnd):
    """取窗口类名（用于辨认 TkChild / TkTopLevel），失败返回空串。"""
    if _user32 is None or not hwnd:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(256)
        _user32.GetClassNameW(ctypes.c_void_p(hwnd), buf, 256)
        return buf.value
    except Exception:
        return ""


def _set_attr(hwnd, attr, value):
    """设置一个 DWM 属性，返回 HRESULT；环境不支持时返回非 0。"""
    if _dwm is None or not hwnd:
        return 1
    v = ctypes.c_int(int(value))
    try:
        return _dwm.DwmSetWindowAttribute(ctypes.c_void_p(hwnd),
                                          ctypes.c_uint(attr),
                                          ctypes.byref(v), ctypes.sizeof(v))
    except Exception:
        return 1


def _get_attr(hwnd, attr):
    """读取一个 DWM 属性（仅供测试用），失败返回 None。"""
    if _dwm is None or not hwnd:
        return None
    v = ctypes.c_int(-1)
    try:
        hr = _dwm.DwmGetWindowAttribute(ctypes.c_void_p(hwnd),
                                        ctypes.c_uint(attr),
                                        ctypes.byref(v), ctypes.sizeof(v))
    except Exception:
        return None
    return v.value if hr == _S_OK else None


def paint(widget, mode="light"):
    """把一个 Tk 顶层窗口（root / Toplevel）的标题栏刷成对应明暗配色。

    mode 传已解析的明暗主题：'light' 或 'dark'（见 colors_for()）。
    标题栏不跟随主题色——底色取应用底 BG_APP，文字取 FG_TEXT，这样系统外框
    和客户区是同一条色系，换主题色时标题栏保持稳定。

    返回 True 表示操作系统接受了 CAPTION_COLOR（Win11 22000+）；
    False 表示仅完成了降级处理（Win10 只切深浅标志）或整体不支持。
    """
    if _dwm is None:
        return False
    hwnd = hwnd_of(widget)
    if not hwnd:
        return False
    spec = colors_for(mode)
    caption = _colorref(spec["caption"])
    text = _colorref(spec["text"])
    border = caption
    # 标题栏底色偏暗时告诉系统按暗色方案渲染字形/按钮 hover，
    # 顺便让 Win10（不支持 35）也能把标题栏切到与主题一致的深浅。
    dark = 1 if _luminance(spec["caption"]) < 0.5 else 0
    hr_dark = _set_attr(hwnd, _DWMWA_USE_IMMERSIVE_DARK_MODE, dark)
    if hr_dark != _S_OK:
        _set_attr(hwnd, _DWMWA_USE_IMMERSIVE_DARK_MODE_OLD, dark)
    _set_attr(hwnd, _DWMWA_CAPTION_COLOR, caption)
    _set_attr(hwnd, _DWMWA_TEXT_COLOR, text)
    _set_attr(hwnd, _DWMWA_BORDER_COLOR, border)
    return _get_attr(hwnd, _DWMWA_CAPTION_COLOR) == caption
