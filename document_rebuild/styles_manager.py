"""
样式管理器 - 创建和管理文档样式
"""

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.shared import qn
from docx.oxml import OxmlElement

def create_document_styles(doc):
    """创建文档样式（使用高级API + 中英文混合字体）"""
    
    print("🎨 创建文档样式...")
    
    styles_created = {}
    
    try:
        # 1. 文档标题样式
        if 'CustomTitle' not in [s.name for s in doc.styles]:
            title_style = doc.styles.add_style('CustomTitle', WD_STYLE_TYPE.PARAGRAPH)
            
            # 字体设置：黑体，二号(22pt)
            font = title_style.font
            font.name = '黑体'
            font.size = Pt(22)
            font.bold = True
            
            # 设置中英文混合字体
            set_mixed_font_for_style(title_style, '黑体', 'Times New Roman')
            
            # 段落设置：居中，行距20磅，段后24磅
            para_format = title_style.paragraph_format
            para_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para_format.line_spacing = Pt(20)
            para_format.space_after = Pt(24)
            
            styles_created['title'] = 'CustomTitle'
            print("  ✅ 标题样式: 黑体22pt，居中")
        
        # 2. 一级标题样式
        if 'CustomHeading1' not in [s.name for s in doc.styles]:
            h1_style = doc.styles.add_style('CustomHeading1', WD_STYLE_TYPE.PARAGRAPH)
            
            # 字体设置：黑体，三号(16pt)
            font = h1_style.font
            font.name = '黑体'
            font.size = Pt(16)
            font.bold = True
            
            # 设置中英文混合字体
            set_mixed_font_for_style(h1_style, '黑体', 'Times New Roman')
            
            # 段落设置：居中，行距20磅
            para_format = h1_style.paragraph_format
            para_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para_format.line_spacing = Pt(20)
            para_format.space_before = Pt(24)
            para_format.space_after = Pt(18)
            
            # 设置大纲级别（使用XML方式）
            style_element = h1_style._element
            ppr = style_element.get_or_add_pPr()
            outline_lvl = OxmlElement('w:outlineLvl')
            outline_lvl.set(qn('w:val'), '0')
            ppr.append(outline_lvl)
            
            styles_created['heading1'] = 'CustomHeading1'
            print("  ✅ 一级标题样式: 黑体16pt，居中，大纲级别1")
        
        # 3. 二级标题样式
        if 'CustomHeading2' not in [s.name for s in doc.styles]:
            h2_style = doc.styles.add_style('CustomHeading2', WD_STYLE_TYPE.PARAGRAPH)
            
            # 字体设置：宋体加粗，小三号(15pt)
            font = h2_style.font
            font.name = '宋体'
            font.size = Pt(15)
            font.bold = True
            
            # 设置中英文混合字体
            set_mixed_font_for_style(h2_style, '宋体', 'Times New Roman')
            
            # 段落设置：左对齐，行距20磅
            para_format = h2_style.paragraph_format
            para_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para_format.line_spacing = Pt(20)
            para_format.space_before = Pt(18)
            para_format.space_after = Pt(12)
            
            # 设置大纲级别（使用XML方式）
            style_element = h2_style._element
            ppr = style_element.get_or_add_pPr()
            outline_lvl = OxmlElement('w:outlineLvl')
            outline_lvl.set(qn('w:val'), '1')
            ppr.append(outline_lvl)
            
            styles_created['heading2'] = 'CustomHeading2'
            print("  ✅ 二级标题样式: 宋体15pt加粗，大纲级别2")
        
        # 4. 三级标题样式
        if 'CustomHeading3' not in [s.name for s in doc.styles]:
            h3_style = doc.styles.add_style('CustomHeading3', WD_STYLE_TYPE.PARAGRAPH)
            
            # 字体设置：宋体，四号(14pt)加粗
            font = h3_style.font
            font.name = '宋体'
            font.size = Pt(14)
            font.bold = True
            
            # 设置中英文混合字体
            set_mixed_font_for_style(h3_style, '宋体', 'Times New Roman')
            
            # 段落设置：缩进2字符，行距20磅
            para_format = h3_style.paragraph_format
            para_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para_format.line_spacing = Pt(20)
            para_format.space_before = Pt(12)
            para_format.space_after = Pt(6)
            para_format.first_line_indent = Pt(28)  # 2字符缩进
            
            # 设置大纲级别（使用XML方式）
            style_element = h3_style._element
            ppr = style_element.get_or_add_pPr()
            outline_lvl = OxmlElement('w:outlineLvl')
            outline_lvl.set(qn('w:val'), '2')
            ppr.append(outline_lvl)
            
            styles_created['heading3'] = 'CustomHeading3'
            print("  ✅ 三级标题样式: 宋体14pt加粗，缩进2字符，大纲级别3")
        
        # 5. 正文样式
        if 'CustomBody' not in [s.name for s in doc.styles]:
            body_style = doc.styles.add_style('CustomBody', WD_STYLE_TYPE.PARAGRAPH)
            
            # 字体设置：宋体，小四号(12pt)
            font = body_style.font
            font.name = '宋体'
            font.size = Pt(12)
            
            # 设置中英文混合字体
            set_mixed_font_for_style(body_style, '宋体', 'Times New Roman')
            
            # 段落设置：行距20磅
            para_format = body_style.paragraph_format
            para_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para_format.line_spacing = Pt(20)
            para_format.space_before = Pt(0)
            para_format.space_after = Pt(0)
            
            styles_created['paragraph'] = 'CustomBody'
            print("  ✅ 正文样式: 宋体12pt，行距20磅，中英文混合字体")
        
        return styles_created
        
    except Exception as e:
        print(f"❌ 创建样式失败: {e}")
        return {}

def set_mixed_font_for_style(style, chinese_font='宋体', english_font='Times New Roman'):
    """为样式设置中英文混合字体"""
    try:
        # 设置基础字体（英文）
        style.font.name = english_font
        
        # 设置东亚字体（中文）
        style_element = style._element
        rpr = style_element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn('w:eastAsia'), chinese_font)
        rfonts.set(qn('w:ascii'), english_font)
        rfonts.set(qn('w:hAnsi'), english_font)
        
        return True
    except Exception as e:
        print(f"    设置混合字体失败: {e}")
        return False

def apply_style_to_paragraph(para, content_type, styles_created, doc):
    """为段落应用样式"""
    try:
        if content_type in styles_created:
            style_name = styles_created[content_type]
            para.style = doc.styles[style_name]
            return True
        else:
            print(f"    未知内容类型: {content_type}")
            return False
    except Exception as e:
        print(f"    应用样式失败: {e}")
        return False