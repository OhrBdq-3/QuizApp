# -*- coding: utf-8 -*-
"""生成界面截图，改 UI 后用它核对实际效果。

用法（在仓库根目录）：python scripts/screenshot.py
输出：outputs/ui_bank.png / ui_quiz.png / ui_answered.png / ui_dark.png / ui_settings.png

会真的弹窗口、抢前台再抓屏，所以别在跑它的时候操作鼠标键盘。
"""
import os
import sys
import time
import tkinter as tk

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import quiz_app.app as qa  # noqa: E402

SAMPLE = [
    qa.Question("示例卷__1", "示例卷", 1, "单选",
                "下列关于人工智能训练师职责的描述，哪一项是正确的？",
                ["A. 只负责数据标注", "B. 参与数据、模型与业务闭环",
                 "C. 只负责模型部署", "D. 与业务无关"], ["B"], "解析：B 最全面。"),
    qa.Question("示例卷__2", "示例卷", 2, "多选",
                "以下哪些属于常见的模型评测指标？",
                ["A. 准确率", "B. 召回率", "C. F1", "D. 代码行数"],
                ["A", "B", "C"], "解析：ABC 是评测指标。"),
    qa.Question("示例卷__3", "示例卷", 3, "判断",
                "微调一定能提升模型在所有任务上的表现。",
                ["A. 正确", "B. 错误"], ["B"], "解析：不一定。"),
    qa.Question("示例卷__4", "示例卷", 4, "单选", "过拟合的典型表现是？",
                ["A. 训练误差低、验证误差高", "B. 训练误差高、验证误差低",
                 "C. 两者都低", "D. 两者都高"], ["A"], "解析：A。"),
    qa.Question("示例卷__5", "示例卷", 5, "单选", "以下哪项不属于数据清洗？",
                ["A. 去重", "B. 缺失值处理", "C. 异常值处理",
                 "D. 增加训练轮数"], ["D"], "解析：D 属于训练策略。"),
]

OUT = os.path.join(ROOT, "outputs")


def grab(win, name):
    """抓窗口的实际屏幕矩形。"""
    from PIL import ImageGrab
    import ctypes
    win.update_idletasks()
    win.update()
    try:
        ctypes.windll.user32.SetForegroundWindow(win.winfo_id())
    except Exception:
        pass
    time.sleep(0.6)
    win.update()
    x, y = win.winfo_rootx(), win.winfo_rooty()
    w, h = win.winfo_width(), win.winfo_height()
    path = os.path.join(OUT, name)
    ImageGrab.grab(bbox=(x, y, x + w, y + h)).save(path)
    print("saved", path, (w, h))


def main():
    os.makedirs(OUT, exist_ok=True)
    root = tk.Tk()
    app = qa.QuizApp(root)
    root.geometry("1300x900+60+30")   # QuizApp 会设默认尺寸，这里覆盖
    app.root.withdraw()
    app.root.update()

    app.db.upsert_sheet("示例卷", "screenshot.xlsx", SAMPLE, seq=0)
    app._load_from_db()
    app._render_bank()
    app.root.deiconify()
    app.root.attributes("-topmost", True)
    app.root.update()
    time.sleep(0.4)
    grab(root, "ui_bank.png")

    app.start_session("示例卷")
    app.root.update()
    time.sleep(0.4)
    grab(root, "ui_quiz.png")

    # 收起 AI 解析栏（persist=False：别把截图脚本的状态写进用户配置）
    app._set_ai_collapsed(True, persist=False)
    app.root.update()
    time.sleep(0.4)
    grab(root, "ui_collapsed.png")
    app._set_ai_collapsed(False, persist=False)
    app.root.update()

    app._click_instant(app._q_at(0), "A")
    app.root.update()
    time.sleep(0.4)
    grab(root, "ui_answered.png")

    app._apply_color_theme("dark", refresh_widgets=True)
    app.root.update()
    time.sleep(0.4)
    grab(root, "ui_dark.png")

    app.open_settings()
    root.update()
    time.sleep(0.5)
    win = app._settings_window
    win.update()
    grab(win, "ui_settings.png")
    root.destroy()


if __name__ == "__main__":
    main()
