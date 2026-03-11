# AI4Word - 开发日志

## 项目概述

本项目创建一个Python应用，接收自然语言指令调用AI生成内容，并在Microsoft Word中以动画效果排版，支持交互式修改。

---

## 开发阶段

### V1.0: 基础AI写作

- 使用 `python-docx` 库创建和保存Word文档
- 调用DeepSeek API生成文本内容
- 文本直接保存为 `.docx` 格式

### V2.0: 实时打字效果

- 切换使用 `pywin32` 与已打开的Word应用交互
- 实现流式接收AI返回数据
- 逐字将内容输入到Word（打字机效果）

### V3.0: Markdown格式化

- 引入 `markdown-it-py` 解析Markdown
- 将格式（标题、粗体、斜体、列表等）正确应用到Word
- 放弃逐字输入以保证格式完整性

### V4.0: 格式化动画写入

- 结合V2和V3的方案
- 先完整解析Markdown并设置格式
- 然后以逐字动画方式写入Word

### V5.0: 交互式排版尝试

- 支持对话式排版指令
- JSON指令化方案（后期被证明局限性大）

### V6.0: AI代码执行

- AI生成可执行的 `pywin32` 代码
- 支持在沙箱环境中动态执行
- 新增自我修复循环机制

### V6.1: 问题修复与文档完善

- 修复 `selection.Style` 赋值的 `AttributeError`
- 改用 `selection.Range.Style = selection.Document.Styles(...)` 设置样式
- 新增 `README.md` 和规范化文档

---

## 核心技术栈

- **AI服务**: DeepSeek API
- **Word交互**: `pywin32` (COM接口)
- **Markdown解析**: `markdown-it-py`
- **API通信**: `requests`
- **环境管理**: `python-dotenv`

---

## 主要文件

- `main.py` - 核心Word排版程序
- `报名.py` - 项目报名提交脚本
- `requirements.txt` - 依赖列表
- `README.md` - 项目说明文档
- `DEVELOPMENT_LOG.md` - 开发文档
