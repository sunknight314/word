"""
大模型分析文档段落的预设Prompt模板
"""

# 预设的段落分析prompt模板
PARAGRAPH_ANALYSIS_PROMPT = """请分析以下Word文档中每个段落的类型和级别，并返回JSON格式的分析结果。

段落内容（每段前20个字符）：
{paragraphs_text}

分类标准：
- title: 文档标题（最高级别）
- heading1: 一级标题（第一章、1.等）
- heading2: 二级标题（1.1、1.2等）
- heading3: 三级标题（1.1.1等）
- heading4: 四级标题或更低级别
- paragraph: 正文段落
- list: 列表项
- quote: 引用文本
- other: 其他类型

识别要点：
- 注意数字编号规律（第一章、1.1、（一）等）
- 标题通常较短
- 中文标题格式：第X章、X.X、（X）等

请返回包含analysis_result数组的JSON对象，每个元素包含paragraph_number、preview_text、type、confidence和reason字段。"""


def format_paragraphs_for_analysis(paragraphs_data: list) -> str:
    """
    格式化段落数据用于大模型分析
    
    Args:
        paragraphs_data: 段落数据列表，每个元素包含paragraph_number和preview_text
    
    Returns:
        使用预设prompt模板格式化的完整prompt
    """
    # 构建段落列表文本（不包含编号，避免干扰分析）
    paragraphs_text = ""
    for para in paragraphs_data:
        paragraphs_text += f"段落{para['paragraph_number']}: {para['preview_text']}\n"
    
    # 使用预设模板
    return PARAGRAPH_ANALYSIS_PROMPT.format(paragraphs_text=paragraphs_text.strip())


def get_analysis_prompt(paragraphs_data: list) -> str:
    """
    获取用于段落分析的完整prompt (旧版本，建议使用get_paragraph_analysis_prompt)
    
    Args:
        paragraphs_data: 段落数据列表
    
    Returns:
        格式化的prompt字符串
    """
    return format_paragraphs_for_analysis(paragraphs_data)


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
- list: 列表项
- quote: 引用文本
- other: 其他类型

识别要点：
1. 注意数字编号规律（第一章、1.1、（一）等）
2. 标题通常较短，正文段落较长
3. 中文标题格式：第X章、X.X、（X）等
4. 根据段落长度和内容判断类型
5. 标题通常包含章节编号或关键词

输出要求：
- 返回完整的JSON对象
- 包含analysis_result数组
- 每个元素包含paragraph_number、preview_text、type、confidence、reason字段
- confidence为0-1之间的数字，表示置信度
- reason简要说明判断依据

JSON格式：
```json
{
  "analysis_result": [
    {
      "paragraph_number": 1,
      "preview_text": "段落预览文本",
      "type": "title/heading1/heading2/heading3/heading4/paragraph/list/quote/other",
      "confidence": 0.95,
      "reason": "判断依据"
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
5. 分析页码格式要求（位置、样式、起始页等）
6. 识别页眉页脚的要求

输出要求：
- 返回完整的JSON配置，确保格式正确
- 不要添加任何解释文字，只返回JSON
- 使用pt（磅）作为基本单位，如"12pt"、"24pt"
- 如果文档中没有明确说明某些格式要求，请使用合理的默认值

常用字号参考：
- 小四12pt、四号14pt、小三15pt、三号16pt、小二18pt、二号22pt、小一24pt、一号26pt

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