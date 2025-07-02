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
    获取用于段落分析的完整prompt
    
    Args:
        paragraphs_data: 段落数据列表
    
    Returns:
        格式化的prompt字符串
    """
    return format_paragraphs_for_analysis(paragraphs_data)