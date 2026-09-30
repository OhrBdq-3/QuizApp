"""Small application-styled modal dialogs for the Tk desktop UI."""
import tkinter as tk
from tkinter import ttk


class AppDialogs:
    def __init__(self, root, font_family="Microsoft YaHei UI"):
        self.root = root
        self.font_family = font_family
        self.palette_provider = None
        self.titlebar_hook = None

    def set_palette_provider(self, provider):
        self.palette_provider = provider

    def set_titlebar_hook(self, hook):
        """注册标题栏染色回调：hook(window)，由宿主按当前明暗主题实现。"""
        self.titlebar_hook = hook

    def showinfo(self, title, message, **kwargs):
        self._show(title, message, kind="info", parent=kwargs.get("parent"))

    def showerror(self, title, message, **kwargs):
        self._show(title, message, kind="error", parent=kwargs.get("parent"))

    def askyesno(self, title, message, **kwargs):
        return bool(self._show(title, message, kind="confirm",
                               parent=kwargs.get("parent")))

    def _show(self, title, message, kind, parent=None):
        parent = parent or self.root
        window = tk.Toplevel(parent)
        window.withdraw()
        window.title(title)
        window.transient(parent)
        window.resizable(False, False)
        palette = self.palette_provider() if self.palette_provider else {}
        bg = palette.get("BG_CARD", "#ffffff")
        fg = palette.get("FG_TEXT", "#202123")
        muted = palette.get("FG_MUTED", "#52525b")
        accent = palette.get("ACCENT", "#242424")
        window.configure(bg=bg)
        result = {"value": False}

        shell = tk.Frame(window, bg=bg, padx=28, pady=24)
        shell.pack(fill="both", expand=True)

        heading = tk.Frame(shell, bg=bg)
        heading.pack(fill="x")
        badge_color = "#dc2626" if kind == "error" else accent
        badge_text = "!" if kind == "error" else ("?" if kind == "confirm" else "i")
        tk.Label(heading, text=badge_text, bg=badge_color, fg="#ffffff",
                 font=(self.font_family, 11, "bold"), width=2, height=1).pack(
                     side="left", anchor="n", padx=(0, 12))
        tk.Label(heading, text=title, bg=bg, fg=fg,
                 font=(self.font_family, 14, "bold"), anchor="w").pack(
                     side="left", fill="x", expand=True)

        tk.Label(shell, text=str(message), bg=bg, fg=muted,
                 font=(self.font_family, 10), justify="left", anchor="w",
                 wraplength=440).pack(fill="x", pady=(18, 24))

        actions = tk.Frame(shell, bg=bg)
        actions.pack(fill="x")

        def close(value=False):
            result["value"] = value
            try:
                window.grab_release()
            except tk.TclError:
                pass
            window.destroy()

        if kind == "confirm":
            ttk.Button(actions, text="取消", style="Soft.TButton",
                       command=lambda: close(False)).pack(side="right")
            ttk.Button(actions, text="确认", style="Accent.TButton",
                       command=lambda: close(True)).pack(side="right", padx=(0, 10))
            window.bind("<Return>", lambda _e: close(True))
        else:
            ttk.Button(actions, text="知道了", style="Accent.TButton",
                       command=lambda: close(True)).pack(side="right")
            window.bind("<Return>", lambda _e: close(True))

        window.bind("<Escape>", lambda _e: close(False))
        window.protocol("WM_DELETE_WINDOW", lambda: close(False))
        window.update_idletasks()
        width = max(460, min(580, shell.winfo_reqwidth() + 10))
        height = max(210, shell.winfo_reqheight() + 10)
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = max(parent.winfo_width(), width)
        parent_h = max(parent.winfo_height(), height)
        x = parent_x + (parent_w - width) // 2
        y = parent_y + (parent_h - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")
        window.deiconify()
        if self.titlebar_hook:
            try:
                self.titlebar_hook(window)
            except Exception:
                pass
        window.lift()
        window.grab_set()
        window.focus_force()
        parent.wait_window(window)
        return result["value"]
