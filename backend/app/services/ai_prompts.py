"""
大模型分析文档段落的预设Prompt模板
"""

def get_paragraph_analysis_prompt(paragraphs_data: list) -> tuple:
    """
    获取段落分析prompt (分为system和user两部分)
    
    Args:
        paragraphs_data: 段落数据列表
        
    Returns:
        (system_prompt, user_prompt) 元组
    """
    system_prompt = """你是一个专业的Word文档段落分析专家。你的任务是分析Word文档中每个段落的类型和级别。

分析标准：
- title: 文档标题（最高级别）
- heading1: 一级标题（第一章、1.等）
- heading2: 二级标题（1.1、1.2等）
- heading3: 三级标题（1.1.1等）
- heading4: 四级标题或更低级别
- paragraph: 正文段落
- abstract_title_cn: 中文摘要标题（摘要、内容摘要等）
- abstract_title_en: 英文摘要标题（Abstract、ABSTRACT等）
- figure_caption: 图注（图1、图1-1、Figure 1等开头）
- table_caption: 表注（表1、表1-1、Table 1等开头）
- quote: 引用文本
- other: 其他类型

识别要点：
1. 注意数字编号规律（第一章、1.1、（一）等）
2. 标题通常较短，正文段落较长
3. 中文标题格式：第X章、X.X、（X）等
4. 中文摘要标题识别：通常为"摘要"、"内容摘要"、"论文摘要"等，一般居中，无编号
5. 英文摘要标题识别：通常为"Abstract"、"ABSTRACT"等，一般居中，无编号
6. 图注识别：以"图"或"Figure"开头，后跟数字编号，如"图1-1 系统架构图"
7. 表注识别：以"表"或"Table"开头，后跟数字编号，如"表2-3 实验数据"
8. 摘要标题通常出现在文档前部，在标题之后、目录之前
9. 图注和表注通常较短，且包含特定的编号格式
10. 根据段落长度和内容判断类型
11. 标题通常包含章节编号或关键词

输出要求：
- 返回完整的JSON对象
- 包含analysis_result数组
- 每个元素包含paragraph_number、type


JSON格式：
```json
{
  "analysis_result": [
    {
      "paragraph_number": 1,
      "type": "title/heading1/heading2/heading3/heading4/paragraph/abstract_title_cn/abstract_title_en/figure_caption/table_caption/quote/other",
    }
  ]
}
```"""
    
    # 构建段落列表文本
    paragraphs_text = ""
    for para in paragraphs_data:
        paragraphs_text += f"段落{para['paragraph_number']}: {para['preview_text']}\n"
    
    user_prompt = f"""请分析以下Word文档中每个段落的类型和级别：

{paragraphs_text.strip()}"""
    
    return system_prompt, user_prompt

def get_format_config_generation_prompt(document_content: str) -> tuple:
    """
    获取格式配置生成prompt (分为system和user两部分)
    
    Args:
        document_content: 文档完整内容
        
    Returns:
        (system_prompt, user_prompt) 元组
    """
    system_prompt = """你是一个专业的Word文档格式分析专家。你的任务是分析Word文档的格式要求，并生成对应的JSON配置文件。

分析要求：
1. 仔细阅读文档中关于字体、字号、行距、对齐方式等格式要求
2. 识别页边距、页面方向、页面大小等页面设置
3. 提取标题层级的格式要求（一级、二级、三级标题的字体大小、加粗、居中等）
4. 识别正文段落的格式要求（首行缩进、行距、字体等）
5. 识别摘要标题的格式要求（中文摘要标题和英文摘要标题的字体、字号、对齐方式等）
6. 特别注意悬挂缩进的要求（如参考文献、引用等可能使用悬挂缩进，第一行顶格，后续行缩进）
7. 分析页码格式要求（位置、样式、起始页等）
8. 识别页眉页脚的要求

输出要求：
- 返回完整的JSON配置，确保格式正确
- 不要添加任何解释文字，只返回JSON
- 直接使用文档中的原始单位，不要进行单位转换
- 保留原始表述，如"2.54厘米"、"1英寸"、"三号"、"16磅"等
- 如果文档中没有明确说明某些格式要求，请使用合理的默认值

单位示例（请保留原始单位）：
- 页边距：如"2.54cm"、"1inch"、"72磅"
- 字号：如"小四"、"三号"、"16pt"、"12磅"
- 行距：如"20磅"、"1.5倍"、"固定值20pt"
- 缩进：如"2字符"、"24pt"、"0.5英寸"

缩进说明：
- first_line_indent: 首行缩进（正值表示首行向右缩进）
- left_indent: 左缩进（整个段落左边距）
- right_indent: 右缩进（整个段落右边距）
- hanging_indent: 悬挂缩进（正值表示除首行外的其他行向右缩进，常用于参考文献、编号列表等）

JSON配置结构：
```json
{{
  "document_info": {{
    "title": "文档标题",
    "author": "作者",
    "description": "描述"
  }},
  "page_settings": {{
    "margins": {{
      "top": "页边距上",
      "bottom": "页边距下", 
      "left": "页边距左",
      "right": "页边距右"
    }},
    "orientation": "portrait/landscape",
    "size": "A4/A3等"
  }},
  "styles": {{
    "title": {{
      "name": "样式名称",
      "font": {{
        "chinese": "中文字体",
        "english": "英文字体",
        "size": "字号(如16pt)",
        "bold": true/false,
        "italic": true/false
      }},
      "paragraph": {{
        "alignment": "left/center/right/justify",
        "line_spacing": "行距(如20pt)",
        "space_before": "段前距(如12pt)",
        "space_after": "段后距(如12pt)",
        "first_line_indent": "首行缩进(如24pt)",
        "left_indent": "左缩进(如0pt)",
        "right_indent": "右缩进(如0pt)",
        "hanging_indent": "悬挂缩进(如0pt)"
      }},
      "outline_level": null或数字
    }},
    "heading1": {{ /* 一级标题样式 */ }},
    "heading2": {{ /* 二级标题样式 */ }},
    "heading3": {{ /* 三级标题样式 */ }},
    "abstract_title_cn": {{
      "name": "ChineseAbstractTitle",
      "font": {{
        "chinese": "中文字体",
        "english": "英文字体",
        "size": "字号(如16pt)",
        "bold": true/false,
        "italic": false
      }},
      "paragraph": {{
        "alignment": "center",
        "line_spacing": "行距",
        "space_before": "段前距",
        "space_after": "段后距",
        "first_line_indent": "0pt",
        "left_indent": "0pt",
        "right_indent": "0pt",
        "hanging_indent": "0pt"
      }},
      "outline_level": null
    }},
    "abstract_title_en": {{
      "name": "EnglishAbstractTitle",
      "font": {{
        "chinese": "中文字体",
        "english": "英文字体",
        "size": "字号(如16pt)",
        "bold": true/false,
        "italic": false
      }},
      "paragraph": {{
        "alignment": "center",
        "line_spacing": "行距",
        "space_before": "段前距",
        "space_after": "段后距",
        "first_line_indent": "0pt",
        "left_indent": "0pt",
        "right_indent": "0pt",
        "hanging_indent": "0pt"
      }},
      "outline_level": null
    }},
    "paragraph": {{ /* 正文样式 */ }}
  }},
  "toc_settings": {{
    "title": "目录标题",
    "levels": "目录层级如1-3",
    "hyperlinks": true/false,
    "tab_leader": "dots/none",
    "page_format": "roman/decimal",
    "page_start": 1,
    "headers": {{
      "odd": "奇数页页眉",
      "even": "偶数页页眉"
    }}
  }},
  "page_numbering": {{
    "toc_section": {{
      "format": "upperRoman/lowerRoman/decimal",
      "start": 起始页码,
      "restart": true/false,
      "template": "页码模板如{{page}}"
    }},
    "content_sections": {{
      "format": "decimal",
      "start": 起始页码,
      "restart_first_chapter": true/false,
      "continue_others": true/false,
      "template": "页码模板如第 {{page}} 页"
    }}
  }},
  "headers_footers": {{
    "odd_even_different": true/false,
    "first_page_different": true/false,
    "toc_section": {{
      "odd_header": "目录奇数页页眉",
      "even_header": "目录偶数页页眉",
      "odd_footer": "目录奇数页页脚",
      "even_footer": "目录偶数页页脚"
    }},
    "content_sections": {{
      "odd_header_template": "正文奇数页页眉模板",
      "even_header_template": "正文偶数页页眉模板",
      "odd_footer": "正文奇数页页脚",
      "even_footer": "正文偶数页页脚"
    }}
  }},
  "section_breaks": {{
    "between_toc_and_content": "odd_page/even_page/new_page/continuous",
    "between_chapters": "odd_page/even_page/new_page/continuous", 
    "between_sections": "continuous/new_page"
  }},
  "document_structure": {{
    "auto_generate_toc": true/false,
    "chapter_start_page": "odd/even/any",
    "chapter_numbering": "arabic/roman",
    "section_numbering": "decimal/roman"
  }}
}}
```

请根据用户提供的文档内容分析格式要求，生成完整的format_config.json配置。"""

    user_prompt = f"""请分析以下Word文档格式要求的文字内容：

{document_content}"""

    return system_prompt, user_prompt