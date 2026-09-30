"""样式预设：一键应用成套排版（论文/公文/简历/博客），替代逐属性硬改。

AI 生成排版代码时逐属性设字体字号既慢又容易顾此失彼；预设把一整套经过
验证的样式定义一次性应用，AI 只需 apply_preset("论文")。每套预设覆盖正文
（字体/字号/首行缩进/行距）与 1-3 级标题（字体/字号/对齐），全部通过修改
样式定义完成，与项目的"格式改样式"黄金法则一致。

预设值参考标准中文文档排版惯例（论文：宋体小四 + 黑体标题 + 1.5 倍行距；
公文：仿宋三号 + 首行缩进两字 + 固定 28 磅行距）。
"""

WD_STYLE_NORMAL = -1
WD_STYLE_HEADING_1 = -2
WD_STYLE_HEADING_2 = -3
WD_STYLE_HEADING_3 = -4
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_LINE_SPACING_SINGLE = 0        # wdLineSpaceSingle
WD_LINE_SPACING_ONE_POINT_FIVE = 1  # wdLineSpace1pt5
WD_LINE_SPACING_DOUBLE = 2       # wdLineSpaceDouble
WD_LINE_SPACING_EXACTLY = 4      # wdLineSpaceExactly

_HEADING_STYLE_IDS = {1: WD_STYLE_HEADING_1, 2: WD_STYLE_HEADING_2, 3: WD_STYLE_HEADING_3}

PRESETS = {
    "论文": {
        "desc": "宋体小四正文、首行缩进两字、1.5 倍行距；黑体标题",
        "body": {"font": "宋体", "size": 12, "first_line_indent": 2,
                 "line_spacing_rule": WD_LINE_SPACING_ONE_POINT_FIVE},
        "headings": {
            1: {"font": "黑体", "size": 22, "align": WD_ALIGN_PARAGRAPH_CENTER},
            2: {"font": "黑体", "size": 16, "align": WD_ALIGN_PARAGRAPH_LEFT},
            3: {"font": "黑体", "size": 15, "align": WD_ALIGN_PARAGRAPH_LEFT},
        },
    },
    "公文": {
        "desc": "仿宋三号正文、首行缩进两字、固定 28 磅行距；黑体标题",
        "body": {"font": "仿宋", "size": 16, "first_line_indent": 2,
                 "line_spacing_rule": WD_LINE_SPACING_EXACTLY, "line_spacing": 28.0},
        "headings": {
            1: {"font": "黑体", "size": 36, "align": WD_ALIGN_PARAGRAPH_CENTER},
            2: {"font": "黑体", "size": 16, "align": WD_ALIGN_PARAGRAPH_LEFT},
            3: {"font": "黑体", "size": 15, "align": WD_ALIGN_PARAGRAPH_LEFT},
        },
    },
    "简历": {
        "desc": "宋体小四正文、无缩进、单倍行距；黑体紧凑标题",
        "body": {"font": "宋体", "size": 12, "first_line_indent": 0,
                 "line_spacing_rule": WD_LINE_SPACING_SINGLE},
        "headings": {
            1: {"font": "黑体", "size": 18, "align": WD_ALIGN_PARAGRAPH_CENTER},
            2: {"font": "黑体", "size": 14, "align": WD_ALIGN_PARAGRAPH_LEFT},
        },
    },
    "博客": {
        "desc": "宋体 11 号正文、无缩进、单倍行距；黑体分级标题",
        "body": {"font": "宋体", "size": 11, "first_line_indent": 0,
                 "line_spacing_rule": WD_LINE_SPACING_SINGLE},
        "headings": {
            1: {"font": "黑体", "size": 22, "align": WD_ALIGN_PARAGRAPH_CENTER},
            2: {"font": "黑体", "size": 18, "align": WD_ALIGN_PARAGRAPH_LEFT},
            3: {"font": "黑体", "size": 14, "align": WD_ALIGN_PARAGRAPH_LEFT},
        },
    },
}


def apply_preset(doc, name):
    """对文档应用一整套样式预设，返回已应用项的摘要。

    通过修改 Normal 与 Heading 1-3 的样式定义实现，正文与标题一次成型；
    行距规则为固定值（公文）时同步设置 LineSpacing 磅值。
    """
    spec = PRESETS.get(name)
    if spec is None:
        raise ValueError(f"未知预设 {name!r}，可用：{preset_names()}")
    applied = []

    body = spec["body"]
    normal = doc.Styles(WD_STYLE_NORMAL)
    normal.Font.Name = body["font"]
    normal.Font.Size = body["size"]
    pf = normal.ParagraphFormat
    pf.LineSpacingRule = body.get("line_spacing_rule", WD_LINE_SPACING_SINGLE)
    if body.get("line_spacing_rule") == WD_LINE_SPACING_EXACTLY:
        pf.LineSpacing = body.get("line_spacing", 28.0)
    indent = body.get("first_line_indent", 0)
    # 首行缩进 N 个字符 = 字号 × N 磅
    pf.FirstLineIndent = normal.Font.Size * indent if indent else 0
    applied.append(f"正文={body['font']}/{body['size']}磅/缩进{indent}字")

    for level, hs in spec.get("headings", {}).items():
        sid = _HEADING_STYLE_IDS.get(level)
        if sid is None:
            continue
        st = doc.Styles(sid)
        st.Font.Name = hs["font"]
        st.Font.Size = hs["size"]
        st.ParagraphFormat.Alignment = hs.get("align", WD_ALIGN_PARAGRAPH_LEFT)
        applied.append(f"H{level}={hs['font']}/{hs['size']}磅")

    return {"preset": name, "applied": applied}


def preset_names():
    """可用预设名称（顿号分隔，给 AI 与用户提示用）。"""
    return "、".join(sorted(PRESETS))


def preset_help():
    """逐条列出预设说明。"""
    return "\n".join(f"- {name}：{spec['desc']}" for name, spec in sorted(PRESETS.items()))