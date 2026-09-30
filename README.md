<div align="center">
  <sub>中文 ｜ <a href="README_EN.md">English</a></sub>
</div>

<div align="center">

<img src="docs/banner.svg" width="780" alt="刷题 —— 本地离线刷题工具，从 Excel 题库开始"/>

![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-5d3a54)
![Dependencies](https://img.shields.io/badge/dependencies-1%20(openpyxl)-8a7968)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-8a7968)
![Offline](https://img.shields.io/badge/offline-100%25-7a9a76)

**导入 Excel 题库，即可开始刷题。**

不联网、不注册。题库、进度、错题都保存在本地电脑上。

</div>

---

## UI

<div align="center">

<img src="docs/screenshots/quiz.png" width="720" alt="做题界面"/><br/>

<br/><br/>

<table>
  <tr>
    <td width="50%" align="center"><img src="docs/screenshots/bank.png" alt="题库页"/><br/></td>
    <td width="50%" align="center"><img src="docs/screenshots/answered.png" alt="判分反馈"/><br/></td>
  </tr>
</table>

</div>

## 简介

离线刷题工具，私人题库不出本机，无题目数量限制，支持自配AI解析。

## 主要功能

| | |
| --- | --- |
| **导入题库** | 支持 `.xlsx`，每个 Sheet 自动变成一张试卷，题库保存在本地 |
| **单选 · 多选 · 判断** | 提交后立即判分，解析与答案一起显示 |
| **错题本** | 答错的题自动收录，重做答对后自动移出 |
| **进度自动保存** | 关闭程序后再次打开，从上次停下的题目继续 |
| **键盘操作** | 可用字母键选答案、方向键翻页 |
| **智能筛题** | 按常见出题规律筛掉只凭关键词就能答对的题，生成一份精简背题库 |
| **界面主题** | 浅色 / 深色两种主题，字号可调 |
| **AI 解析（可选）** | 可接入 OpenAI 兼容接口、Ollama、LM Studio，为当前题目生成解析 |

## 快速开始

**1. 安装依赖。** 整个项目只有一个第三方包，需要 Python 3.9+（推荐 3.11）：

```bash
pip install openpyxl
```

**2. 启动程序。** 界面用的是 Python 自带的 Tkinter，不用额外装 GUI 框架：

```bash
python main.py
# 等价写法：python -m quiz_app
```

**3. 导入题库。** 点「导入题库」，选一份 Excel，开始刷。

<details>
<summary><b>讲究一点：先建虚拟环境</b></summary>
<br>

Windows（PowerShell）：

```powershell
cd D:\path\to\quiz_app
python -m venv .venv
.venv\Scripts\Activate.ps1
# 若提示禁止运行脚本，改用：.venv\Scripts\activate.bat
pip install -r requirements.txt
```

macOS / Linux：

```bash
cd /path/to/quiz_app
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

用 Conda 也行：

```bash
conda create -n quiz python=3.11
conda activate quiz
cd /path/to/quiz_app
pip install -r requirements.txt
```

每次运行前记得先激活环境。

</details>

<details>
<summary><b>启动时报 No module named 'tkinter'？</b></summary>
<br>

说明你的 Python 缺少 Tk 组件：

- **Windows**：重新运行官方安装程序，勾选 **tcl/tk and IDLE**
- **Ubuntu / Debian**：`sudo apt install python3-tk`
- **macOS**：建议用 python.org 官方安装包（Homebrew 版可能缺 Tk）

</details>

## 题库格式

第一行是表头，只有「题目」一列必填，其他列随手加：

| 列名 | 必填 | 说明 |
| --- | --- | --- |
| 题目 | ✅ | 题干 |
| 题型 | 否 | 单选 / 多选 / 判断，默认单选 |
| 选项 | 单选/多选建议填 | 支持写在一个单元格，或拆成「选项A、选项B…」多列 |
| 答案 | 否 | 如 `A`、`A,B`、`对/错`，也可直接写答案文本 |
| 解析 | 否 | 题目解析 |
| 难度 | 否 | 简单 / 中等 / 困难 |

程序对表头写法挺宽容：「题干」「正确答案」这类常见写法都能认，选项标记 `A.` `A、` `A:` 也都行。`samples/示例题库.xlsx` 是一份现成的样例，也可以用 `python scripts/make_sample.py` 重新生成。

## 快捷键

| 按键 | 作用 |
| --- | --- |
| `A` `S` `D` `F` | 选择对应选项（默认值） |
| `Enter` / `Space` | 提交多选题 |
| `←` `↑` | 上一题 |
| `→` `↓` | 下一题 |

答题键能在「设置 · 答题快捷键」里改成任意字母、数字、`Enter` 或 `Space`；方向键翻页是固定绑定。答完题照样能用方向键翻页回看——焦点在解析面板里时，方向键保留滚动，不会误翻页。

中文输入法开着时，字母键会被输入法截走。程序已自动规避，个别环境若仍不生效，切到英文输入即可。

## AI 解析（可选）

不配置完全不影响刷题。

想用的时候，在「设置」里填一个你自己的模型服务：OpenAI、LM Studio、vLLM 选「OpenAI 兼容」，Ollama 选「Ollama」。只有你主动点「AI 解析」，当前这道题才会发出去；密钥只存在本机配置文件里，也可以走 `OPENAI_API_KEY` 环境变量。

## 使用须知

个人学习和题库复习用。别在禁止使用外部工具或 AI 辅助的考试、竞赛、正式考核里用它——这条底线得自己守住。题库内容和 AI 解析都可能有错，拿不准的，自己再核一遍。

## 开源协议

[MIT](LICENSE) —— 可自由使用、修改和分发。

---
