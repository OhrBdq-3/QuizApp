# -*- coding: utf-8 -*-
"""生成一个示例 Excel 题库（samples/示例题库.xlsx），用于测试导入与解析。

运行：python scripts/make_sample.py
"""
import os

import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "计算机基础"

ws.append(["题型", "题目", "选项", "答案", "解析", "难度"])
rows = [
    ("单选", "计算机的基本组成包括 CPU、内存、硬盘和（ ）。",
     "A. 显示器 B. 操作系统 C. 键盘 D. 主板", "D",
     "计算机硬件五大件：运算器、控制器、存储器、输入设备、输出设备。主板是核心硬件。", "简单"),
    ("单选", "1 GB 等于多少 MB？",
     "A. 1024 B. 1000 C. 512 D. 2048", "A",
     "1 GB = 1024 MB。", "简单"),
    ("单选", "以下哪个是操作系统？",
     "A. Word B. Windows C. Excel D. Photoshop", "B",
     "Windows 是操作系统，其余都是应用软件。", "中等"),
    ("判断", "CPU 是计算机的运算核心。",
     "A. 正确 B. 错误", "A",
     "CPU 即中央处理器，负责运算和控制。", "简单"),
    ("判断", "硬盘属于输出设备。",
     "A. 正确 B. 错误", "B",
     "硬盘是存储设备，不属于输出设备。", "简单"),
    ("多选", "以下属于输入设备的有（ ）。",
     "A. 键盘 B. 鼠标 C. 显示器 D. 扫描仪", "A,B,D",
     "键盘、鼠标、扫描仪是输入设备，显示器是输出设备。", "中等"),
    ("多选", "以下属于存储器的有（ ）。",
     "A. 内存 B. 硬盘 C. U盘 D. 打印机", "A,B,C",
     "内存、硬盘、U盘都是存储介质，打印机是输出设备。", "中等"),
]
for r in rows:
    ws.append(list(r))

# 第二个 sheet
ws2 = wb.create_sheet("编程基础")
ws2.append(["题型", "题目", "选项", "答案", "解析", "难度"])
rows2 = [
    ("单选", "Python 中 print(1 + 2) 的输出是？",
     "A. 3 B. 12 C. 1+2 D. 报错", "A",
     "print 输出表达式的计算结果 3。", "简单"),
    ("单选", "以下哪个是合法的 Python 变量名？",
     "A. 2name B. my-var C. _score D. class", "C",
     "变量名可以以 _ 开头，不能以数字开头、不能含减号、不能用关键字。", "中等"),
    ("判断", "Python 是解释型语言。",
     "A. 正确 B. 错误", "A",
     "Python 由解释器逐行执行。", "简单"),
    ("判断", "列表是可变的，元组也是可变的。",
     "A. 正确 B. 错误", "B",
     "元组是不可变的，列表可变。", "中等"),
    ("多选", "Python 的基本数据类型包括（ ）。",
     "A. 整型 B. 字符串 C. 列表 D. 函数", "A,B,C",
     "函数是对象但不是基本数据类型。", "中等"),
    ("多选", "以下哪些是 Python 的内置函数？",
     "A. len() B. print() C. sum() D. sqrt()", "A,B,C",
     "sqrt() 在 math 模块中，不是内置函数。", "中等"),
]
for r in rows2:
    ws2.append(list(r))

# 输出到仓库根的 samples/ 目录（本脚本位于 scripts/ 下）
out = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "samples",
    "示例题库.xlsx",
)
wb.save(out)
print("saved:", out)
