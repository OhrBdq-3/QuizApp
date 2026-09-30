# -*- coding: utf-8 -*-
"""离线刷题 App 启动入口。

    python main.py

逻辑都在 quiz_app/ 包里，这里只负责启动。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from quiz_app.app import main  # noqa: E402

if __name__ == "__main__":
    main()
