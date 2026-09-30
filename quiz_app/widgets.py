# -*- coding: utf-8 -*-
"""可复用的界面组件。

颜色一律由调用方以 palette dict 传入（见 QuizApp._palette），组件本身不引用
app 的全局色常量——这样切换主题/主题色时只要调用 apply_palette() 重画即可，
不用重建控件。

Tk 画不出圆角：Frame / Button / Progressbar / Label 在 clam 主题下都是直角，
ttk style 也没有 radius 选项。所以凡是要圆角的地方都用 Canvas 自己画：

    RoundedFrame  圆角容器（卡片、面板）
    RoundButton   圆角按钮
    RoundProgress 圆角进度条
    Pill          胶囊标签
    Badge         圆角小色标

Canvas 自身背景一律取父容器色，这样圆角外的四个角"透"出父色，视觉上就是圆角。
"""
import re
import tkinter as tk
from tkinter import font as tkfont

from . import theme

# 选项文本前缀：'A.' 'B、' 'C:' 'D）' 等，徽标已经显示了字母，正文里去掉更干净
_OPT_PREFIX_RE = re.compile(r"^\s*[A-Ha-h]\s*[.、:：)）]\s*")

# 语义色上的文字（绿/红底固定用白字，浅色深色主题都成立）
_ON_SEMANTIC = "#ffffff"


def strip_option_prefix(option: str) -> str:
    """去掉选项文本开头的字母标记，徽标已经承担了字母的显示。"""
    text = _OPT_PREFIX_RE.sub("", option or "", count=1)
    return text or (option or "")


def _bg_of(widget, depth=4):
    """取容器背景色：圆角控件外侧的四个角要"透"出父容器色。

    ttk 控件的背景在 style 里，cget("bg") 会抛错或返回空，所以逐级往上找。
    """
    w = widget
    for _ in range(depth):
        if w is None:
            break
        try:
            color = str(w.cget("bg"))
            if color:
                return color
        except (tk.TclError, AttributeError):
            pass
        w = getattr(w, "master", None)
    return "#ffffff"


def _font_obj(widget, font):
    """把 (family, size, weight) 变成可测量的 Font 对象。"""
    try:
        return tkfont.Font(root=widget, font=font) if font else tkfont.Font(root=widget)
    except (tk.TclError, TypeError):
        return tkfont.Font(root=widget)


class RoundedFrame(tk.Frame):
    """圆角容器：Canvas 画圆角底 + 描边，真实内容一律放进 self.body。

    用法：
        card = RoundedFrame(parent, radius=theme.R_LG, bg=BG_CARD, border=BORDER)
        card.pack(fill="x")
        tk.Label(card.body, ...).pack(...)      # 注意是 .body

    换主题时：本对象的 bg / highlightbackground 走 config() 即可重画（内部已
    接管），外圈底色由 apply_palette() 从父容器重新取。
    """

    def __init__(self, master, *, radius=theme.R_MD, bg="#ffffff",
                 border="#e5e7eb", outer=None, padx=0, pady=0, shadow=0, **kw):
        border = kw.pop("highlightbackground", None) or border
        kw.pop("highlightthickness", None)
        kw.pop("highlightcolor", None)
        bg = kw.pop("background", None) or kw.pop("bg", None) or bg
        self._bg = bg
        self._border = border
        self._radius = radius
        self._padx = padx
        self._pady = pady
        self._shadow = max(0, int(shadow or 0))
        outer_bg = outer or _bg_of(master)
        super().__init__(master, bg=outer_bg, **kw)

        self._canvas = tk.Canvas(self, highlightthickness=0, bd=0,
                                 bg=outer_bg, takefocus=0)
        self._canvas.pack(fill="both", expand=True)
        # body 用 place 而不是 canvas.create_window：内嵌窗口的"请求尺寸"会被
        # 画布 item 的 -width/-height 覆盖成上一轮的实际尺寸，读它来定尺寸就成
        # 了自反馈——容器只会越撑越宽，外层按 weight 分列宽彻底失效。place 属于
        # 外部几何管理，只改实际尺寸、不污染请求尺寸，body.winfo_req* 才是内容
        # 真正的自然尺寸。（具体位置与大小每次 _redraw 时重算）
        self.body = tk.Frame(self._canvas, bg=bg)
        self.body.place(x=padx, y=pady)
        # 投影：Tk 没有真阴影，这里在卡片下方垫一块"底色压深一档"的同形状
        # 圆角块，露出 sw 像素，视觉上就是卡片从纸面上抬起来一点。必须先于
        # 本体创建，才能被本体盖住上半部分。
        self._shadow_item = None
        if self._shadow:
            self._shadow_item = theme.rounded_rect_poly(
                self._canvas, 1, 1, 200, 200, r=radius, fill=outer_bg,
                outline=outer_bg)
        # 用 *_poly：尺寸每次布局都要重设，圆角图元必须是多边形
        self._shape = theme.rounded_rect_poly(self._canvas, 1, 1, 200, 200,
                                              r=radius, fill=bg, outline=border)
        self.body.bind("<Configure>", self._on_body_resize)
        self._canvas.bind("<Configure>", self._on_canvas_resize)
        self._paint_shadow()

    # ---------- 尺寸同步 ----------

    def _inset(self):
        """内容窗口相对画布的内缩量。

        内容是一个矩形子窗口，如果铺满画布，四个直角会把画好的圆角盖掉。
        圆角在 45° 方向离角 r*(1-1/√2)≈0.29r，内缩这么多内容就完全落在
        圆角之内了（多留 1px 让描边露出来）。
        """
        if self._radius <= 0:
            return 0
        return self._radius * 0.293 + 1

    def _on_body_resize(self, event=None):
        """内容自然尺寸变化时，把请求尺寸往上传递（fill="x" 的卡片靠这个定高）。

        这里只能报"内容要多大"，不能跟当前实际尺寸取最大值：画布一旦被父容器
        拉宽过一次，请求宽度就被钉死在那，再也缩不回来（棘轮效应）。外层按
        weight 分配列宽会因此完全失效——AI 解析栏怎么调都收不窄就是栽在这。
        """
        try:
            off = self._inset()
            rw = self.body.winfo_reqwidth() + (self._padx + off) * 2
            rh = self.body.winfo_reqheight() + (self._pady + off) * 2 \
                + self._shadow      # 投影要额外占一条高度，否则会被裁掉
            self._canvas.configure(width=rw, height=rh)
        except tk.TclError:
            pass
        self._redraw()

    def _on_canvas_resize(self, event=None):
        self._redraw()

    def _paint_shadow(self):
        """投影色由父容器底色现场派生，这样换主题不用额外传色。"""
        if not self._shadow or self._shadow_item is None:
            return
        try:
            outer = str(self._canvas.cget("bg"))
            color = theme.shadow_of(outer)
            self._canvas.itemconfig(self._shadow_item, fill=color,
                                    outline=color)
        except tk.TclError:
            pass

    def _redraw(self):
        """按实际尺寸重画圆角底（含投影），并把内容窗口放进圆角之内。"""
        try:
            off = self._inset()
            sw = self._shadow
            cw = self._canvas.winfo_width()
            ch = self._canvas.winfo_height()
            rw = self.body.winfo_reqwidth() + (self._padx + off) * 2
            rh = self.body.winfo_reqheight() + (self._pady + off) * 2 + sw
            w = max(cw if cw > 1 else rw, 8)
            h = max(ch if ch > 1 else rh, 8)
            self.body.place_configure(
                x=self._padx + off, y=self._pady + off,
                width=max(w - (self._padx + off) * 2, 1),
                height=max(h - sw - (self._pady + off) * 2, 1))
            if self._shadow_item is not None:
                theme.rounded_rect_item(self._canvas, self._shadow_item,
                                        2, 1 + sw, w - 2, h - 1,
                                        r=self._radius)
            theme.rounded_rect_item(self._canvas, self._shape, 1, 1,
                                    w - 1, max(h - 1 - sw, 4),
                                    r=self._radius)
            self._paint_shadow()
        except tk.TclError:
            pass

    # ---------- 配色 ----------

    def set_colors(self, bg=None, border=None, outer=None):
        if bg is not None:
            self._bg = bg
        if border is not None:
            self._border = border
        try:
            if outer is not None:
                tk.Frame.configure(self, bg=outer)
                self._canvas.config(bg=outer)
            self.body.config(bg=self._bg)
            self._canvas.itemconfig(self._shape, fill=self._bg,
                                    outline=self._border)
            self._paint_shadow()
        except tk.TclError:
            pass

    def apply_palette(self, palette):
        """换主题：底色/描边由 config 递归更新，这里补外圈底色。

        必须显式重设一次 *self._bg / _border 对应的填充：Canvas 的 bg 是画布
        底色、图元 fill 是独立项，_recolor_widget_tree 改 bg 时图元纹丝不动
        （已验证：canvas.configure(bg=…) 不改 create_polygon 的 fill）。所以
        圆角块永远停在初始色上，透明/继承式传色在这里会变成"写死色"。
        """
        try:
            self.set_colors(outer=_bg_of(self.master))
        except tk.TclError:
            pass

    # ---------- 让 bg / highlightbackground 语义落在圆角底上 ----------

    def configure(self, cnf=None, **kw):
        opts = {}
        if isinstance(cnf, dict):
            opts.update(cnf)
        opts.update(kw)
        if not opts:
            return super().configure()
        handled = {}
        for key in ("bg", "background", "highlightbackground", "highlightcolor"):
            if key in opts:
                handled[key] = opts.pop(key)
        if handled:
            self.set_colors(
                bg=handled.get("bg", handled.get("background")),
                border=handled.get("highlightbackground"))
        if opts:
            return super().configure(**opts)

    config = configure

    def cget(self, key):
        k = str(key).lstrip("-")
        if k in ("bg", "background"):
            return self._bg
        if k in ("highlightbackground", "highlightcolor"):
            return self._border
        return super().cget(key)

    def __getitem__(self, key):
        return self.cget(key)


class OptionRow(RoundedFrame):
    """一个选项行：圆形字母徽标 + 选项文本，整行可点击。

    单选/判断内嵌 Radiobutton、多选内嵌 Checkbutton（都设 indicatoron=False，
    外观就是普通一行文字），这样保留了 Tk 原生的变量语义与键盘可达性，
    只是把排版换成了市面刷题 App 常见的「徽标 + 文字」卡片式行。
    """

    BADGE = 30          # 徽标直径（正方形画布边长）

    def __init__(self, master, letter, text, *, multi=False, variable=None,
                 value=None, command=None, palette=None, font=None):
        pal = dict(palette or {})
        super().__init__(master, radius=theme.R_MD,
                         bg=pal.get("bg", "#ffffff"),
                         border=pal.get("border", "#e5e7eb"))
        self.letter = letter
        self.palette = pal
        self.font = font
        self._multi = multi
        self._variable = variable
        self._value = value
        self._result = None        # None | 'correct' | 'wrong' | 'muted'
        self._hover = False
        self._answered = False

        # 先建按钮：它是 body.winfo_children()[0]，外部（含测试）按位置取更稳定
        if multi:
            self.button = tk.Checkbutton(
                self.body, text=text, variable=variable, font=font,
                indicatoron=False, relief="flat", offrelief="flat", borderwidth=0,
                anchor="w", justify="left", highlightthickness=0, padx=0, pady=0)
        else:
            self.button = tk.Radiobutton(
                self.body, text=text, value=value, variable=variable, command=command,
                font=font, indicatoron=False, relief="flat", offrelief="flat",
                borderwidth=0, anchor="w", justify="left",
                highlightthickness=0, padx=0, pady=0)
        self.button.pack(side="left", fill="x", expand=True,
                         padx=(theme.SP_1, theme.SP_3), pady=theme.SP_4)

        # 徽标：真圆角靠 Canvas。用 before= 插到按钮左侧——视觉上是
        # 「徽标 + 文字」，但 children 顺序里按钮仍排第一。
        self.badge = tk.Canvas(self.body, width=self.BADGE, height=self.BADGE,
                               highlightthickness=0, bd=0, bg=pal.get("bg", "#ffffff"))
        self.badge.pack(side="left", before=self.button, padx=(theme.SP_4, 0))
        pad = 2
        self._badge_shape = theme.rounded_rect(
            self.badge, pad, pad, self.BADGE - pad, self.BADGE - pad,
            r=self.BADGE / 2, fill=pal.get("btn_soft", "#eaecf0"),
            outline=pal.get("btn_soft", "#eaecf0"))
        self._badge_text = self.badge.create_text(
            self.BADGE / 2, self.BADGE / 2, text=letter,
            font=self._badge_font(), fill=pal.get("muted", "#707078"))

        # Tk 子控件的事件不会冒泡到父 Frame；背景、留白和徽标也要绑定。
        # 文字按钮仍交给原生绑定处理，避免多选一次点击切换两次。
        for w in (self, self._canvas, self.body, self.badge, self.button):
            w.bind("<Button-1>", self._on_click)
            w.bind("<Enter>", lambda e: self._set_hover(True))
            w.bind("<Leave>", lambda e: self._set_hover(False))
        if multi and variable is not None:
            variable.trace_add("write", lambda *_: self._repaint())
        self._repaint()

    def _badge_font(self):
        """徽标字号随正文缩放，但比正文小一号、加粗"""
        fam, size, _w = self.font if isinstance(self.font, (tuple, list)) and len(self.font) == 3 \
            else (None, theme.FS_BODY, "normal")
        return (fam, max(8, int(size) - 1), "bold")

    # ---------- 状态 ----------

    def _is_selected(self) -> bool:
        if self._variable is None:
            return False
        try:
            return bool(self._variable.get())
        except (tk.TclError, TypeError):
            return False

    def _set_hover(self, on: bool):
        if self._answered or self._hover == on:
            return
        self._hover = on
        self._repaint()

    def _on_click(self, event):
        if event.widget is not self.button:
            # 点在行的空白处：转交给内部按钮，保持单选/多选的原生行为
            if not self._answered:
                self.button.invoke()
            return
        # 点在内嵌按钮上：单选会立即判分并重绘，多选靠变量 trace 重绘
        if not self._answered and self._multi:
            self._repaint()

    def set_answered(self, answered: bool):
        """作答后锁住整行，并去掉悬停高亮"""
        self._answered = answered
        if answered:
            self._hover = False
        self.button.config(state="disabled" if answered else "normal")

    def apply_result(self, state):
        """state: None / 'correct' / 'wrong' / 'muted'"""
        self._result = state
        self._repaint()

    def apply_palette(self, palette):
        self.palette = dict(palette or {})
        try:
            self.set_colors(outer=_bg_of(self.master))
        except tk.TclError:
            pass
        self._repaint()

    # ---------- 绘制 ----------

    def _repaint(self):
        pal = self.palette
        res = self._result
        if res == "correct":
            row_bg, border = pal["ok_soft"], pal["ok"]
            b_bg, b_fg, txt_fg = pal["ok"], _ON_SEMANTIC, pal["ok"]
        elif res == "wrong":
            row_bg, border = pal["bad_soft"], pal["bad"]
            b_bg, b_fg, txt_fg = pal["bad"], _ON_SEMANTIC, pal["bad"]
        elif res == "muted":
            row_bg, border = pal["bg"], pal["border"]
            b_bg, b_fg, txt_fg = pal["subtle"], pal["faint"], pal["muted"]
        elif self._is_selected():
            row_bg, border = pal["accent_soft"], pal["accent"]
            b_bg, b_fg, txt_fg = pal["accent"], pal["on_accent"], pal["text"]
        else:
            row_bg = pal["subtle"] if self._hover else pal["bg"]
            border = pal["accent"] if self._hover else pal["border"]
            b_bg = pal["btn_soft_hover"] if self._hover else pal["btn_soft"]
            b_fg, txt_fg = pal["muted"], pal["text"]
        try:
            self.set_colors(bg=row_bg, border=border)
            self.button.config(bg=row_bg, fg=txt_fg, activebackground=row_bg,
                               activeforeground=txt_fg, selectcolor=row_bg,
                               disabledforeground=txt_fg)
            self.badge.config(bg=row_bg)
            self.badge.itemconfig(self._badge_shape, fill=b_bg, outline=b_bg)
            self.badge.itemconfig(self._badge_text, fill=b_fg)
        except tk.TclError:
            pass

    def winfo_children(self):
        """对外暴露内容层的子控件（画布是背景，不算内容）。"""
        return self.body.winfo_children()


class Pill(tk.Canvas):
    """胶囊标签：题型、难度、进度状态这类短元信息（两端全圆）。"""

    def __init__(self, master, text, palette, *, fg=None, bg=None,
                 font=None, padx=10, pady=4, radius=None):
        self._fg = fg or palette.get("accent")
        self._bg = bg or palette.get("accent_soft")
        self._text = text or ""
        self._font = font
        self._padx = padx
        f = _font_obj(master, font)
        w = max(int(f.measure(self._text)) + padx * 2, padx * 2 + 4)
        h = int(f.metrics("linespace")) + pady * 2
        super().__init__(master, width=w, height=h, highlightthickness=0, bd=0,
                         bg=_bg_of(master))
        self._radius = radius if radius is not None else h / 2
        # 胶囊两端就是半圆，椭圆图元最贴合；高度不变，宽度变化时改 4 个点即可
        self._shape = self.create_oval(1, 1, w - 1, h - 1,
                                       fill=self._bg, outline=self._bg)
        self._text_id = self.create_text(w / 2, h / 2, text=self._text,
                                         fill=self._fg, font=font)

    def restyle(self, fg=None, bg=None):
        if fg:
            self._fg = fg
        if bg:
            self._bg = bg
        try:
            self.itemconfig(self._shape, fill=self._bg, outline=self._bg)
            self.itemconfig(self._text_id, fill=self._fg)
        except tk.TclError:
            pass

    def config(self, cnf=None, **kw):
        opts = dict(cnf) if isinstance(cnf, dict) else {}
        opts.update(kw)
        if not opts:
            # 无参调用是"取全部选项"，必须透传（换主题时的颜色扫描靠这个）。
            # Canvas 没有 foreground 选项，但胶囊的文字色也要跟着主题走，
            # 所以这里补一个，让重着色逻辑能扫到它（cget 已一并接管）。
            info = super().configure()
            info["foreground"] = ("foreground", "foreground", "Foreground",
                                  self._fg, self._fg)
            return info
        if "text" in opts:
            self._set_text(opts.pop("text"))
        if "bg" in opts or "background" in opts:
            bg = opts.pop("bg", None) or opts.pop("background", None)
            if bg:
                self._bg = bg
                self.restyle(bg=self._bg)
        if "fg" in opts or "foreground" in opts:
            fg = opts.pop("fg", None) or opts.pop("foreground", None)
            if fg:
                self._fg = fg
                self.restyle(fg=self._fg)
        if opts:
            return super().config(**opts)

    configure = config

    def _set_text(self, text):
        self._text = text or ""
        try:
            f = _font_obj(self, self._font)
            w = max(int(f.measure(self._text)) + self._padx * 2, self._padx * 2 + 4)
            self.itemconfig(self._text_id, text=self._text)
            self.configure(width=w)
            h = int(self.cget("height"))
            self.coords(self._shape, 1, 1, w - 1, h - 1)
            self.coords(self._text_id, w / 2, h / 2)
        except tk.TclError:
            pass

    def cget(self, key):
        k = str(key).lstrip("-")
        if k == "text":
            return self._text
        if k in ("bg", "background"):
            return self._bg
        if k in ("fg", "foreground"):
            return self._fg
        return super().cget(key)

    def apply_palette(self, palette):
        """换主题：外层底色从父容器重取（走 Canvas 原生 config，别碰胶囊填充色）。"""
        try:
            tk.Canvas.configure(self, bg=_bg_of(self.master))
        except tk.TclError:
            pass


class Badge(tk.Canvas):
    """圆角小色块，用于答题卡图例之类的色标。"""

    def __init__(self, master, color, size=12, radius=4):
        super().__init__(master, width=size, height=size,
                         highlightthickness=0, bd=0, bg=_bg_of(master))
        self._size = size
        self._radius = radius
        self._shape = theme.rounded_rect(self, 0, 0, size, size, radius,
                                         fill=color, outline=color)

    def recolor(self, color):
        self.itemconfig(self._shape, fill=color, outline=color)


# 按钮变体 → 取色键。hover/禁用态按同一套规则派生，避免每个调用点各写一遍。
_BUTTON_VARIANTS = {
    #                  底             悬停底             文字          加粗
    "soft":       ("btn_soft",        "btn_soft_hover",  "text",      False),
    "accent":     ("accent",          "accent_hover",    "on_accent", True),
    "danger":     ("btn_soft",        "bad_soft",        "muted",     False),
    # 侧栏当前项：做成"浮在侧栏上的一张暖白卡片 + 主色文字"，比浅底块更有层次
    "side":       ("card",            "card",            "accent",    True),
    "nav_soft":   ("btn_soft",        "btn_soft_hover",  "text",      False),
    "nav_accent": ("accent",          "accent_hover",    "on_accent", True),
    "ghost":      ("bg",              "subtle",          "muted",     False),
}

_DEFAULT_PAD = {"danger": (10, 6), "nav_soft": (18, 11), "nav_accent": (18, 11)}


class RoundButton(tk.Canvas):
    """圆角按钮。

    ttk 的 clam 主题没有圆角选项，这里用 Canvas 自己画。接口兼容 ttk.Button
    的常用子集：text / command / state / width（字符数）/ config() / invoke()。
    """

    def __init__(self, master, text="", command=None, variant="soft",
                 width=None, font=None, palette=None, radius=theme.R_SM + 4,
                 padx=None, pady=None, outer=None):
        self._text = text or ""
        self._command = command
        self._variant = variant if variant in _BUTTON_VARIANTS else "soft"
        self._font = font
        self._palette = dict(palette or {})
        self._radius = radius
        self._state = "normal"
        self._hover = False
        self._pressed = False
        self._width_chars = width
        self._padx, self._pady = padx, pady
        if padx is None or pady is None:
            dpad = _DEFAULT_PAD.get(self._variant, (16, 9))
            self._padx = padx if padx is not None else dpad[0]
            self._pady = pady if pady is not None else dpad[1]
        w, h = self._measure(master)
        super().__init__(master, width=w, height=h, highlightthickness=0, bd=0,
                         bg=outer or _bg_of(master), takefocus=0)
        self._shape = theme.rounded_rect_poly(self, 1, 1, w - 1, h - 1, r=radius,
                                              fill="#cccccc", outline="#cccccc")
        self._text_id = self.create_text(w / 2, h / 2, text=self._text,
                                         fill="#000000", font=self._font)
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self._draw()

    # ---------- 尺寸 ----------

    def _measure(self, master):
        try:
            f = _font_obj(master, self._font)
            tw = f.measure(self._text)
            if self._width_chars:
                tw = max(tw, f.measure("0") * abs(int(self._width_chars)))
            lh = f.metrics("linespace")
        except tk.TclError:
            tw, lh = len(self._text) * 7, 16
        return int(tw + self._padx * 2 + 2), int(lh + self._pady * 2 + 2)

    def _relayout(self):
        w, h = self._measure(self)
        try:
            self.configure(width=w, height=h)
        except tk.TclError:
            pass
        self._draw()

    # ---------- 状态与绘制 ----------

    def _colors(self):
        pal = self._palette
        if self._state == "disabled":
            return pal.get("subtle", "#f1f2f4"), pal.get("faint", "#9ca3af")
        bg_key, hover_key, fg_key, _bold = _BUTTON_VARIANTS[self._variant]
        if self._hover or self._pressed:
            bg = pal.get(hover_key, pal.get(bg_key, "#e5e7eb"))
            fg = pal.get(fg_key, "#202123")
            if self._variant == "danger":
                fg = pal.get("bad", fg)
        else:
            bg = pal.get(bg_key, "#e5e7eb")
            fg = pal.get(fg_key, "#202123")
        return bg, fg

    def _draw(self):
        try:
            # 实际尺寸优先：父容器空间不足时按钮会被压扁，此时若仍按请求高度
            # 居中，文字就跑到框外了。winfo_* 还是 1（尚未布局）才退回请求尺寸。
            ww, hh = self.winfo_width(), self.winfo_height()
            w = max(ww if ww > 1 else int(self.cget("width")), 8)
            h = max(hh if hh > 1 else int(self.cget("height")), 8)
            bg, fg = self._colors()
            theme.rounded_rect_item(self, self._shape, 1, 1, w - 1, h - 1,
                                    r=self._radius)
            self.itemconfig(self._shape, fill=bg, outline=bg)
            self.itemconfig(self._text_id, text=self._text, fill=fg)
            self.coords(self._text_id, w / 2, h / 2)
        except tk.TclError:
            pass

    def _on_enter(self, event):
        if self._state == "disabled":
            return
        self._hover = True
        try:
            self.configure(cursor="hand2")
        except tk.TclError:
            pass
        self._draw()

    def _on_leave(self, event):
        self._hover = False
        self._pressed = False
        self._draw()

    def _on_press(self, event):
        if self._state == "disabled":
            return
        self._pressed = True
        self._draw()

    def _on_release(self, event):
        if self._state == "disabled":
            return
        self._pressed = False
        self._draw()
        try:
            x, y = event.x, event.y
            if 0 <= x <= self.winfo_width() and 0 <= y <= self.winfo_height():
                self.invoke()
        except tk.TclError:
            pass

    def invoke(self):
        if self._state == "disabled":
            return None
        if callable(self._command):
            return self._command()
        return None

    # ---------- 兼容 ttk.Button 的接口 ----------

    def config(self, cnf=None, **kw):
        opts = dict(cnf) if isinstance(cnf, dict) else {}
        opts.update(kw)
        if not opts:
            # 无参调用是"取全部选项"，必须透传（换主题时的颜色扫描靠这个）
            return super().configure()
        if "text" in opts:
            self._text = opts.pop("text") or ""
            self._relayout()
        if "command" in opts:
            self._command = opts.pop("command")
        if "state" in opts:
            self._state = str(opts.pop("state"))
            if self._state == "disabled":
                self._hover = False
                self._pressed = False
            self._draw()
        if "font" in opts:
            self._font = opts.pop("font")
            self.itemconfig(self._text_id, font=self._font)
            self._relayout()
        if "palette" in opts:
            self._palette = dict(opts.pop("palette") or {})
            self._draw()
        if opts:
            return super().config(**opts)

    configure = config

    def cget(self, key):
        k = str(key).lstrip("-")
        if k == "state":
            return self._state
        if k == "text":
            return self._text
        return super().cget(key)

    def __getitem__(self, key):
        return self.cget(key)

    def apply_palette(self, palette):
        self._palette = dict(palette or {})
        try:
            tk.Canvas.configure(self, bg=_bg_of(self.master))
        except tk.TclError:
            pass
        self._draw()


class RoundProgress(tk.Canvas):
    """圆角进度条（ttk.Progressbar 在 clam 下是直角的细条）。"""

    def __init__(self, master, value=0, maximum=100, variable=None,
                 length=260, thickness=8, palette=None, outer=None):
        self._value = value
        self._maximum = max(float(maximum or 100), 1e-6)
        self._thickness = thickness
        self._palette = dict(palette or {})
        self._variable = variable
        super().__init__(master, width=length, height=thickness,
                         highlightthickness=0, bd=0,
                         bg=outer or _bg_of(master), takefocus=0)
        self._trough = theme.rounded_rect_poly(self, 0, 0, length, thickness,
                                               r=thickness / 2, fill="#e5e7eb",
                                               outline="#e5e7eb")
        self._bar = theme.rounded_rect_poly(self, 0, 0, length, thickness,
                                            r=thickness / 2, fill="#242424",
                                            outline="#242424")
        self.bind("<Configure>", lambda e: self._draw())
        if variable is not None:
            try:
                variable.trace_add("write", lambda *_: self._sync_var())
            except (tk.TclError, AttributeError):
                pass
        self._draw()

    def _sync_var(self):
        try:
            self._value = float(self._variable.get())
        except (tk.TclError, TypeError, ValueError):
            return
        self._draw()

    def _draw(self):
        try:
            ww, hh = self.winfo_width(), self.winfo_height()
            w = max(ww if ww > 1 else int(self.cget("width")), 4)
            h = max(hh if hh > 1 else int(self.cget("height")), 4)
            trough = self._palette.get("track", self._palette.get("border", "#e5e7eb"))
            fill = self._palette.get("accent", "#242424")
            pct = max(0.0, min(1.0, self._value / self._maximum))
            theme.rounded_rect_item(self, self._trough, 0, 0, w, h, r=h / 2)
            self.itemconfig(self._trough, fill=trough, outline=trough)
            bw = max(h, w * pct)
            theme.rounded_rect_item(self, self._bar, 0, 0, bw, h, r=h / 2)
            self.itemconfig(self._bar, fill=fill, outline=fill)
        except tk.TclError:
            pass

    def config(self, cnf=None, **kw):
        opts = dict(cnf) if isinstance(cnf, dict) else {}
        opts.update(kw)
        if not opts:
            # 无参调用是"取全部选项"，必须透传（换主题时的颜色扫描靠这个）
            return super().configure()
        if "value" in opts:
            self._value = float(opts.pop("value") or 0)
            self._draw()
        if "maximum" in opts:
            self._maximum = max(float(opts.pop("maximum") or 100), 1e-6)
            self._draw()
        if "palette" in opts:
            self._palette = dict(opts.pop("palette") or {})
            self._draw()
        if opts:
            return super().config(**opts)

    configure = config

    def cget(self, key):
        k = str(key).lstrip("-")
        if k == "value":
            return self._value
        if k == "maximum":
            return self._maximum
        return super().cget(key)

    def apply_palette(self, palette):
        self._palette = dict(palette or {})
        try:
            tk.Canvas.configure(self, bg=_bg_of(self.master))
        except tk.TclError:
            pass
        self._draw()
