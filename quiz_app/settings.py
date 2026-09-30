"""综合设置：统一读写 app_settings.json，并提供主题/字体预设。"""
import json
import os

DEFAULTS = {
    # 外观
    "font_scale": 1.0,          # 0.8 - 1.5
    "theme": "system",          # system / light / dark
    "accent": "plum",           # 主题色预设 id，见 ACCENT_THEMES
    # 答题
    "auto_next_delay": 700,     # 答对后自动下一题的延迟（毫秒）
    "default_order": "random",  # 默认顺序：random / original / type
    # 界面
    "ai_panel_collapsed": False,  # AI 解析栏是否收起（收起后题目区铺满）
}

THEME_LABELS = {"system": "跟随系统", "light": "浅色", "dark": "深色"}

# ---------- 主题色（accent）预设 ----------
# 每个预设给出 5 个派生色，保证与深色/浅色背景都协调：
#   ACCENT         主色（按钮、选中、进度条）
#   ACCENT_HOVER   悬停
#   ACCENT_ACTIVE  按下
#   ACCENT_SOFT    浅底（选中行背景、高亮块）
#   FG_ON_ACCENT   主色上的文字（白/黑）
#
# 前面四组是「暖纸质」主色系（与整体界面同源：深梅紫 / 陶土橙 / 鼠尾草绿 /
# 沙金），后面是常见的冷色预设。
ACCENT_THEMES = {
    "plum": {  # 深梅紫（默认，暖色主色）
        "label": "暖梅紫",
        "ACCENT": "#4d3045", "ACCENT_HOVER": "#5e3b53",
        "ACCENT_ACTIVE": "#3a2433", "ACCENT_SOFT": "#e8dfe5",
        "FG_ON_ACCENT": "#ffffff",
    },
    "terra": {  # 陶土橙（暖色强调）
        "label": "陶土橙",
        "ACCENT": "#b65f42", "ACCENT_HOVER": "#c56e50",
        "ACCENT_ACTIVE": "#96482f", "ACCENT_SOFT": "#f3dfd6",
        "FG_ON_ACCENT": "#ffffff",
    },
    "sage": {  # 鼠尾草绿
        "label": "鼠尾草绿",
        "ACCENT": "#627a67", "ACCENT_HOVER": "#718a75",
        "ACCENT_ACTIVE": "#4c6251", "ACCENT_SOFT": "#dfe8df",
        "FG_ON_ACCENT": "#ffffff",
    },
    "sand": {  # 沙金
        "label": "沙金褐",
        "ACCENT": "#a57c45", "ACCENT_HOVER": "#b58b52",
        "ACCENT_ACTIVE": "#87632f", "ACCENT_SOFT": "#eee2ce",
        "FG_ON_ACCENT": "#ffffff",
    },
    "midnight": {  # 深空黑（Windows 深色默认）
        "label": "深空黑",
        "ACCENT": "#242424", "ACCENT_HOVER": "#3a3a3a",
        "ACCENT_ACTIVE": "#111111", "ACCENT_SOFT": "#ececed",
        "FG_ON_ACCENT": "#ffffff",
    },
    "win11": {  # Windows 11 默认强调色（蓝）
        "label": "Windows 蓝",
        "ACCENT": "#0067C0", "ACCENT_HOVER": "#0055A0",
        "ACCENT_ACTIVE": "#003A6E", "ACCENT_SOFT": "#E1F0FA",
        "FG_ON_ACCENT": "#ffffff",
    },
    "vscode": {  # VS Code 编辑器蓝
        "label": "VS Code 蓝",
        "ACCENT": "#007ACC", "ACCENT_HOVER": "#0066A8",
        "ACCENT_ACTIVE": "#004D7A", "ACCENT_SOFT": "#E3F2FD",
        "FG_ON_ACCENT": "#ffffff",
    },
    "violet": {  # VS Code 紫（Visual Studio 品牌色）
        "label": "VS Code 紫",
        "ACCENT": "#7C4DFF", "ACCENT_HOVER": "#6940D6",
        "ACCENT_ACTIVE": "#4A2EB0", "ACCENT_SOFT": "#F0E9FF",
        "FG_ON_ACCENT": "#ffffff",
    },
    "teal": {  # 深青（沉稳，浅色/深色都好看）
        "label": "青碧",
        "ACCENT": "#0E7C86", "ACCENT_HOVER": "#0B656E",
        "ACCENT_ACTIVE": "#074A51", "ACCENT_SOFT": "#E0F3F5",
        "FG_ON_ACCENT": "#ffffff",
    },
    "emerald": {  # 翡翠绿
        "label": "翡翠绿",
        "ACCENT": "#16A34A", "ACCENT_HOVER": "#15803D",
        "ACCENT_ACTIVE": "#166534", "ACCENT_SOFT": "#E4F5E9",
        "FG_ON_ACCENT": "#ffffff",
    },
    "amber": {  # 琥珀橙（暖色）
        "label": "琥珀橙",
        "ACCENT": "#EA580C", "ACCENT_HOVER": "#C2410C",
        "ACCENT_ACTIVE": "#9A3412", "ACCENT_SOFT": "#FDEEDD",
        "FG_ON_ACCENT": "#ffffff",
    },
    "rose": {  # 玫瑰红
        "label": "玫瑰红",
        "ACCENT": "#E11D48", "ACCENT_HOVER": "#BE123C",
        "ACCENT_ACTIVE": "#881337", "ACCENT_SOFT": "#FCE7EC",
        "FG_ON_ACCENT": "#ffffff",
    },
}

ACCENT_THEME_IDS = list(ACCENT_THEMES.keys())

# 三层表面：BG_APP（页面底）→ BG_CARD（卡片）→ BG_SUBTLE（卡片内的次级面板）。
# 关键区别是 BG_APP 必须比 BG_CARD 深一档，否则白卡片浮在白底上毫无层次，
# 看起来就像没做设计。深色模式同理，只是方向反过来（卡片比底色亮一档）。
#
# 整体走「暖纸质」：底色是暖米纸，卡片是暖白，描边/文字都是偏褐的暖灰，
# 不是纯白 + 中性灰。BG_SIDE 是侧栏（比页面底再深一档），SHADOW 是卡片
# 投影（Tk 没有真阴影，用一条比底色更深的圆角块模拟）。
THEME_PALETTES = {
    "light": {
        "BG_APP": "#f5f0e8", "BG_CARD": "#fffdfa", "BG_SUBTLE": "#faf5ec",
        "BG_SIDE": "#ede6dc", "SHADOW": "#e9dfd0",
        "BORDER": "#ded5c8", "BORDER_SOFT": "#e9e0d2",
        "FG_TEXT": "#29251f", "FG_MUTED": "#746d63", "FG_FAINT": "#9a9186",
        "ACCENT": "#4d3045", "ACCENT_HOVER": "#5e3b53",
        "ACCENT_ACTIVE": "#3a2433", "ACCENT_SOFT": "#e8dfe5",
        "FG_ON_ACCENT": "#ffffff", "C_UNANSWERED": "#e7dfd2",
        "C_TEXT_ON_LIGHT": "#4a433a", "C_OK_SOFT": "#dfe8df",
        "C_BAD_SOFT": "#f5ded2", "C_OK": "#477358", "C_BAD": "#b44d43",
        "BTN_SOFT": "#ece5da", "BTN_SOFT_HOVER": "#e2d9cb",
    },
    "dark": {
        "BG_APP": "#14120e", "BG_CARD": "#1d1a15", "BG_SUBTLE": "#241f19",
        "BG_SIDE": "#191610", "SHADOW": "#000000",
        "BORDER": "#332c23", "BORDER_SOFT": "#2b251d",
        "FG_TEXT": "#ece5d8", "FG_MUTED": "#a2988a", "FG_FAINT": "#7b7263",
        "ACCENT": "#4d3045", "ACCENT_HOVER": "#5e3b53",
        "ACCENT_ACTIVE": "#3a2433", "ACCENT_SOFT": "#e8dfe5",
        "FG_ON_ACCENT": "#ffffff", "C_UNANSWERED": "#2e2822",
        "C_TEXT_ON_LIGHT": "#16130f", "C_OK_SOFT": "#22322a",
        "C_BAD_SOFT": "#3a2420", "C_OK": "#7ea88c", "C_BAD": "#d99086",
        "BTN_SOFT": "#2b251e", "BTN_SOFT_HOVER": "#372f26",
    },
}

# 顺序模式 -> 顶栏下拉框文案
ORDER_LABELS = {
    "random": "随机乱序",
    "original": "原顺序",
    "type": "题型分组",
}


def clamp(v, lo, hi):
    try:
        return max(lo, min(hi, float(v)))
    except (TypeError, ValueError):
        return lo


def normalize(cfg):
    """把任意 dict 收敛到合法范围，缺失字段用默认值补齐。"""
    out = dict(DEFAULTS)
    if isinstance(cfg, dict):
        for k in DEFAULTS:
            if k in cfg:
                out[k] = cfg[k]
    out["font_scale"] = clamp(out.get("font_scale", 1.0), 0.8, 1.5)
    out["auto_next_delay"] = int(clamp(out.get("auto_next_delay", 700), 100, 3000))
    out["default_order"] = out.get("default_order", "random")
    if out["default_order"] not in ORDER_LABELS:
        out["default_order"] = "random"
    theme = str(out.get("theme", "system")).lower()
    out["theme"] = theme if theme in THEME_LABELS else "system"
    accent = str(out.get("accent", "midnight"))
    out["accent"] = accent if accent in ACCENT_THEMES else "midnight"
    out["ai_panel_collapsed"] = bool(out.get("ai_panel_collapsed", False))
    return out


def system_theme():
    """Return the Windows app theme; other systems safely default to light."""
    try:
        import winreg
        path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return "light" if value else "dark"
    except (ImportError, OSError):
        return "light"


def resolve_theme(theme):
    return system_theme() if theme == "system" else theme


def accent_colors(accent_id, resolved_theme=None):
    """返回某个 accent 预设的 5 个颜色 dict（拷贝，可安全修改）。

    深色主题下，近黑的主色（如默认的深空黑 #242424）会和深色背景糊在一起，
    按钮看起来"消失"。所以这里对深色主题做一次反向提亮：主色变浅、其上的
    文字变深，保持"高对比的强调块"这个语义不变。
    """
    theme = ACCENT_THEMES.get(accent_id, ACCENT_THEMES["midnight"])
    colors = {k: v for k, v in theme.items() if k != "label"}
    if resolved_theme == "dark" and _luminance(colors["ACCENT"]) < 0.35:
        # 混白而不是直接 _lighten：_lighten 会把色相洗掉（深梅紫会变灰），
        # 混白能保住原本的暖色调。
        base = _mix(colors["ACCENT"], "#ffffff", 0.58)
        colors = {
            "ACCENT": base,
            "ACCENT_HOVER": _shade(base, 0.88),      # 更亮
            "ACCENT_ACTIVE": _shade(base, 0.98),
            "ACCENT_SOFT": _mix(base, THEME_PALETTES["dark"]["BG_CARD"], 0.82),
            "FG_ON_ACCENT": "#1b1d21",
        }
    return colors


def load_config(path):
    cfg = {}
    try:
        with open(path, encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, ValueError, TypeError):
        pass
    return normalize(cfg)


def save_config(path, cfg):
    cfg = normalize(cfg)
    with open(path + ".tmp", "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(path + ".tmp", path)
    return cfg


# ---------- 自定义主题色派生 ----------

def _shade(hex_color, factor):
    """factor>1 变暗，factor<1 变亮。hex 形如 #rrggbb。"""
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    if factor >= 1:
        r = int(r * (2 - factor))
        g = int(g * (2 - factor))
        b = int(b * (2 - factor))
    else:
        r = int(r + (255 - r) * (1 - factor))
        g = int(g + (255 - g) * (1 - factor))
        b = int(b + (255 - b) * (1 - factor))
    r, g, b = (max(0, min(255, c)) for c in (r, g, b))
    return f"#{r:02x}{g:02x}{b:02x}"


def _lighten(hex_color, t):
    """朝白色混合 t (0..1)。"""
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    r = int(r + (255 - r) * t)
    g = int(g + (255 - g) * t)
    b = int(b + (255 - b) * t)
    return f"#{max(0, min(255, r)):02x}{max(0, min(255, g)):02x}{max(0, min(255, b)):02x}"


def _mix(color_a, color_b, t):
    """按 t (0..1) 把 color_a 混向 color_b。"""
    def parts(c):
        h = c.lstrip("#")
        if len(h) == 3:
            h = h[0] * 2 + h[1] * 2 + h[2] * 2
        return [int(h[i:i + 2], 16) for i in (0, 2, 4)]

    a, b = parts(color_a), parts(color_b)
    out = [int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3)]
    return f"#{out[0]:02x}{out[1]:02x}{out[2]:02x}"


def _luminance(hex_color):
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


def luminance(hex_color):
    """亮度（0=黑 1=白），供外部判断色块上的文字色。"""
    return _luminance(hex_color)


def derive_accent(hex_color):
    """由任意 hex 主色派生 hover/active/soft/on_accent 四个变体（自定义色用）。"""
    accent = (hex_color or "#242424").strip() or "#242424"
    on = "#ffffff" if _luminance(accent) < 0.6 else "#202123"
    return {
        "ACCENT": accent,
        "ACCENT_HOVER": _shade(accent, 1.15),
        "ACCENT_ACTIVE": _shade(accent, 1.35),
        "ACCENT_SOFT": _lighten(accent, 0.88),
        "FG_ON_ACCENT": on,
    }


def font_scale_factor(scale):
    return clamp(scale, 0.8, 1.5)
