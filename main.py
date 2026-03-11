import os
import time
import requests
import win32com.client
from dotenv import load_dotenv
from markdown_it import MarkdownIt

# --- Word COM 常量 ---
WD_STYLE_NORMAL = -1 # 普通段落样式
WD_STYLE_HEADING_1 = -2 # 一级标题样式
WD_STYLE_HEADING_2 = -3 # 二级标题样式
WD_STYLE_HEADING_3 = -4 # 三级标题样式
WD_STYLE_LIST_BULLET = -49 # 无序列表样式
WD_ALIGN_PARAGRAPH_LEFT = 0 # 左对齐
WD_ALIGN_PARAGRAPH_CENTER = 1 # 居中对齐
WD_ALIGN_PARAGRAPH_RIGHT = 2 # 右对齐
WD_REPLACE_ALL = 2 # 替换所有
WD_PRINT_VIEW = 3 # 打印视图
WD_ORIENT_LANDSCAPE = 1 # 横向纸张方向
WD_LINE_SPACING_1_5 = 1 # 1.5倍行间距

def get_word():
    """获取或创建一个Word应用实例。"""
    try:
        app = win32com.client.GetObject(None, "Word.Application")
        print("已连接到现有的Word程序。" )
    except Exception:
        try:
            app = win32com.client.Dispatch("Word.Application")
            print("启动了一个新的Word程序。" )
        except Exception: return None
    app.Visible = True
    return app

def ai_request(prompt: str, api_key: str, system_prompt: str):
    """通用的AI调用函数。"""
    url = "https://api.deepseek.com/chat/completions"
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    data = {"model": "deepseek-chat", "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": prompt}]}
    try:
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        return response.json()['choices'][0]['message']['content']
    except Exception as e:
        print(f"调用API时出错: {e}")
        return None

def write(sel, markdown_text: str):
    """将Markdown文本解析并以动画形式写入Word。"""
    md = MarkdownIt() # 初始化Markdown解析器
    tokens = md.parse(markdown_text)
    style_map = {1: WD_STYLE_HEADING_1, 2: WD_STYLE_HEADING_2, 3: WD_STYLE_HEADING_3} # 标题样式映射
    is_first_block, is_in_list = True, False # 控制段落和列表状态
    for token in tokens:
        if token.type in ['paragraph_open', 'heading_open', 'bullet_list_open']: # 处理段落、标题和列表开始
            if not is_first_block: sel.TypeParagraph()
            is_first_block = False
        if token.type == 'heading_open':
            style_id = style_map.get(int(token.tag[1]), WD_STYLE_NORMAL)
            sel.Range.Style = sel.Document.Styles(style_id)
        elif token.type == 'paragraph_open' and not is_in_list:
            sel.Range.Style = sel.Document.Styles(WD_STYLE_NORMAL)
        elif token.type == 'bullet_list_open': is_in_list = True
        elif token.type == 'list_item_open': sel.Range.Style = sel.Document.Styles(WD_STYLE_LIST_BULLET)
        elif token.type == 'inline' and token.children:
            for child in token.children:
                if child.type == 'text':
                    for char in child.content:
                        sel.TypeText(char); time.sleep(0.01)
                elif child.type == 'strong_open': sel.Font.Bold = True
                elif child.type == 'strong_close': sel.Font.Bold = False
                elif child.type == 'em_open': sel.Font.Italic = True
                elif child.type == 'em_close': sel.Font.Italic = False
        elif token.type == 'bullet_list_close':
            is_in_list = False; sel.TypeParagraph(); sel.Range.Style = sel.Document.Styles(WD_STYLE_NORMAL); is_first_block = True

def gen_code(prompt: str, api_key: str):
    """(初次尝试) 让AI生成用于Word自动化的Python代码。"""
    system_prompt = f"""
你是世界一流的Python程序员，专精于使用 `pywin32` 库对 Microsoft Word 文档进行自动化排版。
用户的指令是关于如何格式化一个已经打开的Word文档。你的任务是编写一小段Python代码来实现这个排版需求。

**可用的全局变量**:
你可以在代码中直接使用两个已经为你定义好的全局变量：
- `word_app`: Word应用程序对象。
- `doc`: 当前活动的文档对象。

**三大黄金法则**:
1.  **永远直接修改样式 (Style)**: 对于任何关于字体、段落格式（如对齐、缩进、行距）的修改，**必须**通过修改文档的样式定义来完成。这是最可靠、最高效的方法。
    -   **示例**: `doc.Styles(WD_STYLE_NORMAL).Font.Name = "宋体"`
    -   **反例 (不要这样做)**: 不要用 `Find` 或遍历所有段落的方式来修改格式。

2.  **精确计算缩进**: 当用户要求“首行缩进N个字符”时，你必须基于该样式的字号来计算缩进的磅值。
    -   **正确示例 (缩进2字符)**:
        ```python
        style = doc.Styles(WD_STYLE_NORMAL)
        # 1个字符的宽度约等于其字号大小（磅值）
        indent_points = style.Font.Size * 2
        style.ParagraphFormat.FirstLineIndent = indent_points
        ```

3.  **极简的错误处理**:
    -   **禁止**对绝大多数操作（如修改样式、字体、段落）使用 `try...except`。让主程序去捕获错误。
    -   **唯一的例外**: 只有在设置页面边距 (`doc.PageSetup`) 时，才需要用 `try...except Exception as e: pass` 包裹，因为它有时会不稳定。

**输出要求**:
- **只输出纯Python代码**。不要包含 `import` 语句、函数定义、`word_app`或`doc`的重新声明，也不要添加任何Markdown标记 (如 ```python) 或解释性文字。

**常量与字号参考**:
- 样式: `WD_STYLE_NORMAL`(-1), `WD_STYLE_HEADING_1`(-2), `WD_STYLE_HEADING_2`(-3), `WD_STYLE_HEADING_3`(-4)
- 对齐: `WD_ALIGN_PARAGRAPH_LEFT`(0), `WD_ALIGN_PARAGRAPH_CENTER`(1), `WD_ALIGN_PARAGRAPH_RIGHT`(2)
- 视图: `WD_PRINT_VIEW`(3)
- 字号 (磅值): "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18, "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9
"""
    print("正在请求AI生成自动化代码（初次尝试）...")
    code_str = ai_request(prompt, api_key, system_prompt)
    if not code_str: return None
    clean_code = code_str.strip().replace("```python", "").replace("```", "").strip()
    print("AI生成的代码:\n---\n" + clean_code + "\n---")
    return clean_code

def fix_code(prompt: str, failed_code: str, error_msg: str, api_key: str):
    """(自我修复) 当代码执行失败后，让AI根据错误信息生成一段修正代码。"""
    system_prompt = f"""
你是一个顶级的Python调试专家，专精于`pywin32`库的Word自动化。
我之前为了完成用户的排版指令：“{prompt}”，运行了你编写的如下代码：
--- FAILED CODE ---
{failed_code}
--- END FAILED CODE ---

然而，这段代码执行后失败了，具体的错误信息是：
`{error_msg}`

你的任务是：
1. 仔细分析失败的代码和错误信息。
2. 思考一种**全新的、不同的方法**来完成用户的原始指令。不要只是对之前的代码做微小的修改。
3. 编写一段新的Python代码片段来解决这个问题。例如，如果`Find/Replace`失败了，可以尝试用`For Each`循环遍历段落 `for p in doc.Paragraphs:` 再进行操作。

请只输出修正后的、完整的Python代码片段，并且同样要遵守**不要自己处理异常**的核心规则（PageSetup除外）。
"""
    print("\n代码执行失败。正在请求AI进行自我修复...")
    code_str = ai_request(prompt, api_key, system_prompt)
    if not code_str: return None
    clean_code = code_str.strip().replace("```python", "").replace("```", "").strip()
    print("AI生成的修正代码:\n---\n" + clean_code + "\n---")
    return clean_code

def main():
    """主函数，采用“AI自我修复”的方案。"""
    dotenv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    load_dotenv(dotenv_path=dotenv_path)
    api_key = os.getenv("DEEPSEEK_API_KEY")
    if not api_key or api_key == "YOUR_DEEPSEEK_API_KEY_HERE":
        print("错误：未能加载有效的 DEEPSEEK_API_KEY。" ); return

    app = get_word()
    if not app: print("错误：无法启动或连接到Word。" ); return
    doc = app.Documents.Add() if app.Documents.Count == 0 else app.ActiveDocument
    sel = app.Selection
    print("\nWord已准备就绪。" )

    # --- 阶段一：初始内容生成 ---
    prompt = input("请输入您的写作需求：")
    if prompt:
        system_prompt = "你是一个乐于助人的助手，你总是使用Markdown格式进行回复。"
        resp = ai_request(prompt, api_key, system_prompt)
        if resp:
            print("\n正在格式化并以动画效果写入Word...")
            write(sel, resp)
            print("写入完成。" )

    # --- 阶段二：对话式排版 ---
    print("\n--- 现在进入AI智能排版模式（支持自我修复）---")
    print("您可以输入复杂的排版指令，AI将为您编写并调试代码。" )
    print("输入 '退出', 'exit' 或 'quit' 来结束程序。" )
    
    exec_globals = {"word_app": app, "doc": doc, "WD_STYLE_NORMAL":-1, "WD_STYLE_HEADING_1":-2, "WD_STYLE_HEADING_2":-3, "WD_STYLE_HEADING_3":-4, "WD_STYLE_LIST_BULLET":-49, "WD_ALIGN_PARAGRAPH_LEFT":0, "WD_ALIGN_PARAGRAPH_CENTER":1, "WD_ALIGN_PARAGRAPH_RIGHT":2, "WD_REPLACE_ALL":2, "WD_PRINT_VIEW":3, "WD_ORIENT_LANDSCAPE":1, "WD_LINE_SPACING_1_5":1}

    while True:
        fmt = input("\n您想如何调整格式？> ")
        if not fmt: continue
        if fmt.lower() in ['退出', 'exit', 'quit']: print("程序已结束。"); break

        # 1. 初次尝试
        code_to_exec = gen_code(fmt, api_key)
        if code_to_exec:
            try:
                print("正在执行AI生成的代码（第1次尝试）...")
                exec(code_to_exec, exec_globals)
                print("代码执行完毕，格式已应用。" )
                continue 
            except Exception as e:
                print(f"执行初次代码时发生错误: {e}")
                
                # 2. 自我修复
                corrected_code = fix_code(fmt, code_to_exec, str(e), api_key)
                if corrected_code:
                    try:
                        print("正在执行AI生成的修正代码（第2次尝试）...")
                        exec(corrected_code, exec_globals)
                        print("修正代码执行完毕，格式已应用。" )
                    except Exception as e_corrected:
                        print(f"执行修正代码时再次发生错误: {e_corrected}")
                        print("AI自我修复失败，请尝试一个更简单的指令。" )
                else:
                    print("AI未能生成有效的修正代码。" )
        else:
            print("无法理解您的指令或AI未返回有效代码。" )

if __name__ == "__main__":
    main()