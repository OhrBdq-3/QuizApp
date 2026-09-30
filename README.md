# 离线刷题 App

**中文** | [English](README_EN.md)

一个简单好用的本地刷题工具：导入 Excel 题库，马上开始刷题。

不需要联网、不需要注册，所有数据都保存在你自己的电脑上。

## 它能做什么

- **导入即刷题** —— 支持 `.xlsx` 题库，每个 Sheet 自动变成一张试卷，导入一次永久保存
- **三种题型** —— 单选题、多选题、判断题，答完立即判分并显示解析
- **灵活刷题** —— 随机乱序 / 原始顺序 / 按题型分组，随时切换
- **错题本** —— 答错的题自动收录，重做答对后自动移出
- **进度自动保存** —— 关掉程序再打开，从上次做到的地方继续
- **智能筛题** —— 按常见出题规律筛掉"背关键词就行"的题目，生成精简背题库
- **AI 解析（可选）** —— 接入 OpenAI 兼容接口、Ollama、LM Studio 等模型服务，一键生成考点分析和易错点提示

## 安装

需要 **Python 3.9 或更高版本**（推荐 3.11）。界面使用 Python 自带的 Tkinter，无需额外安装 GUI 框架。

建议用虚拟环境安装依赖，避免污染系统 Python。下面两种方式选一种即可。

### 方式一：venv

Windows（PowerShell）：

```powershell
# 1. 进入项目目录
cd D:\人工智能平台工作\工具\quiz_app

# 2. 创建虚拟环境
python -m venv .venv

# 3. 激活虚拟环境
.venv\Scripts\Activate.ps1
# 若提示禁止运行脚本，改用：.venv\Scripts\activate.bat

# 4. 安装依赖
pip install -r requirements.txt
```

macOS / Linux：

```bash
cd /path/to/quiz_app
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 方式二：Conda

```bash
# 1. 创建并激活环境
conda create -n quiz python=3.11
conda activate quiz

# 2. 进入项目目录后安装依赖
cd /path/to/quiz_app
pip install -r requirements.txt
```

### 只装一个依赖也可以

如果不想用虚拟环境，直接装唯一的第三方依赖也行：

```bash
pip install openpyxl
```

## 启动

每次运行前记得先激活环境（venv：`.venv\Scripts\Activate.ps1`；conda：`conda activate quiz`），然后：

```bash
python main.py
# 等价写法：python -m quiz_app
```

> **提示：** 如果启动时报 `No module named 'tkinter'`，说明你的 Python 缺少 Tk 组件：
> - Windows：重新运行官方安装程序，勾选 **"tcl/tk and IDLE"**
> - Ubuntu/Debian：`sudo apt install python3-tk`
> - macOS：建议使用 python.org 官方安装包（Homebrew 版可能缺失）

## 题库格式

Excel 第一行为表头，只需一列 **「题目」** 即可导入，其余列都是可选的：

| 列名 | 必填 | 说明 |
| --- | --- | --- |
| 题目 | ✅ | 题干 |
| 题型 | 否 | 单选 / 多选 / 判断，默认单选 |
| 选项 | 单选/多选建议填 | 支持写在一个单元格，或拆成「选项A、选项B…」多列 |
| 答案 | 否 | 如 `A`、`A,B`、`对/错`，也可直接写答案文本 |
| 解析 | 否 | 题目解析 |
| 难度 | 否 | 简单 / 中等 / 困难 |

程序会自动兼容常见表头写法（如「题干」「正确答案」）和多种选项标记（`A.` `A、` `A:` 等）。`samples/示例题库.xlsx` 可直接参考，也可用 `python scripts/make_sample.py` 重新生成。


## 关于 AI 解析

AI 功能完全可选，不配置不影响刷题。只有主动点击「AI 解析」时，才会把当前题目发送给你自己配置的模型服务。

## 使用须知

本项目用于个人学习和题库复习。请勿在禁止使用外部工具或 AI 辅助的考试、竞赛及正式考核中使用，使用者应自行遵守相关规则。题库内容和 AI 解析可能存在错误，请自行核实。

## 开源协议

本项目基于 [MIT License](LICENSE) 发布，可自由使用、修改和分发。
