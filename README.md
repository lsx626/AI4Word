# AI4Word

> 一个通过调用DeepSeek，实现对Word内容生成和格式编辑的Agent项目。运行在终端，无GUI。用户可一次性生成内容，并多次调整排版（包括但不限于页边距、行间距、字体大小和样式、纸张方向等），直到达到满意效果。

## 功能概述

- 接收自然语言写作需求并调用 DeepSeek AI 生成内容。
- 解析 Markdown 并将格式（标题、粗体、斜体、列表等）应用到 Word。
- 以逐字动画（打字机效果）将文本写入活动的 Word 文档。
- 支持交互式排版：AI 可生成并执行 pywin32 代码以实现复杂排版需求（页边距、行间距、字体大小、纸张方向等），并可在失败后尝试自我修复。

## 前置条件

- Windows 操作系统（需安装 Microsoft Word）。
- Python 3.8+ 环境（建议使用虚拟环境或 conda）。
- 在项目根目录准备 `.env`，包含：

```env
DEEPSEEK_API_KEY=你的_api_key
```

## 安装

1. 创建并激活虚拟环境（可选）：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

1. 安装依赖：

```bash
pip install -r requirements.txt
```

## 使用方法

```bash
python main.py
```

交互说明：

- 程序启动后会提示输入写作需求；AI 返回 Markdown 文本，程序会以动画效果写入活动的 Word 文档。
- 写入完成后程序会进入"AI 智能排版模式"，你可以输入排版指令（例如"将所有一级标题居中并改为蓝色"），程序会请求 AI 生成用于 `pywin32` 的代码并尝试执行。

## 常见问题与排查

- 报错 `AttributeError: Property '<unknown>.Style' can not be set.`
  - 说明：这是因为不能直接对 `Selection.Style` 赋值。代码已修复为通过 `selection.Range.Style = selection.Document.Styles(...)` 来设置样式。
  - 解决：请确保使用的是本仓库最新版本的 `main.py`，并在 Word 中允许外部程序访问（Word 的 Trust Center 设置）。

- 如果程序无法连接到 Word：确认 Word 已安装且可以通过 COM 被脚本控制；检查防火墙或 UAC 设置。

## 开发者说明

- 主程序文件：`main.py`（位于 `pre` 目录）。
- 开发日志：`DEVELOPMENT_LOG.md`，记录了版本演进与主要修复。

## 许可证

未指定特定许可证。
