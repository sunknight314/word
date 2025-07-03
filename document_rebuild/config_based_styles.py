"""
基于配置的样式管理器 - 从JSON配置创建文档样式
"""

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.shared import qn
from config_loader import FormatConfigLoader

def create_styles_from_config(doc, config_loader):
    """基于配置创建文档样式"""
    
    print("🎨 基于配置创建文档样式...")
    
    styles_config = config_loader.get_styles_config()
    styles_created = {}
    
    try:
        for content_type, style_config in styles_config.items():
            style_name = style_config.get("name", f"Custom{content_type.title()}")
            
            # 检查样式是否已存在
            if style_name not in [s.name for s in doc.styles]:
                style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                
                # 应用字体设置
                apply_font_settings(style, style_config.get("font", {}), config_loader)
                
                # 应用段落设置
                apply_paragraph_settings(style, style_config.get("paragraph", {}), config_loader)
                
                # 设置大纲级别
                outline_level = style_config.get("outline_level")
                if outline_level is not None:
                    set_outline_level(style, outline_level)
                
                styles_created[content_type] = style_name
                print(f"  ✅ {content_type}样式: {style_name}")
            else:
                styles_created[content_type] = style_name
                print(f"  ♻️ {content_type}样式已存在: {style_name}")
        
        return styles_created
        
    except Exception as e:
        print(f"❌ 创建样式失败: {e}")
        return {}

def apply_font_settings(style, font_config, config_loader):
    """应用字体设置"""
    try:
        font = style.font
        
        # 基本字体设置
        chinese_font = font_config.get("chinese", "宋体")
        english_font = font_config.get("english", "Times New Roman")
        font_size = font_config.get("size", "12pt")
        
        font.name = english_font
        font.size = config_loader.parse_length(font_size)
        font.bold = font_config.get("bold", False)
        font.italic = font_config.get("italic", False)
        
        # 设置中英文混合字体
        set_mixed_font_for_style(style, chinese_font, english_font)
        
        return True
    except Exception as e:
        print(f"    设置字体失败: {e}")
        return False

def apply_paragraph_settings(style, para_config, config_loader):
    """应用段落设置"""
    try:
        para_format = style.paragraph_format
        
        # 对齐方式
        alignment = para_config.get("alignment", "left")
        para_format.alignment = config_loader.parse_alignment(alignment)
        
        # 行距
        line_spacing = para_config.get("line_spacing", "20pt")
        if line_spacing:
            para_format.line_spacing = config_loader.parse_length(line_spacing)
        
        # 段前距
        space_before = para_config.get("space_before", "0pt")
        if space_before:
            para_format.space_before = config_loader.parse_length(space_before)
        
        # 段后距  
        space_after = para_config.get("space_after", "0pt")
        if space_after:
            para_format.space_after = config_loader.parse_length(space_after)
        
        # 首行缩进
        first_line_indent = para_config.get("first_line_indent", "0pt")
        if first_line_indent and first_line_indent != "0pt":
            para_format.first_line_indent = config_loader.parse_length(first_line_indent)
        
        # 左缩进
        left_indent = para_config.get("left_indent", "0pt")
        if left_indent and left_indent != "0pt":
            para_format.left_indent = config_loader.parse_length(left_indent)
        
        # 右缩进
        right_indent = para_config.get("right_indent", "0pt")
        if right_indent and right_indent != "0pt":
            para_format.right_indent = config_loader.parse_length(right_indent)
        
        # 悬挂缩进
        hanging_indent = para_config.get("hanging_indent", "0pt")
        if hanging_indent and hanging_indent != "0pt":
            # 悬挂缩进是负的首行缩进
            para_format.first_line_indent = -config_loader.parse_length(hanging_indent)
        
        return True
    except Exception as e:
        print(f"    设置段落格式失败: {e}")
        return False

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

def set_outline_level(style, level):
    """设置大纲级别"""
    try:
        style_element = style._element
        ppr = style_element.get_or_add_pPr()
        outline_lvl = OxmlElement('w:outlineLvl')
        outline_lvl.set(qn('w:val'), str(level))
        ppr.append(outline_lvl)
        return True
    except Exception as e:
        print(f"    设置大纲级别失败: {e}")
        return False

def apply_page_settings(doc, config_loader):
    """应用页面设置"""
    try:
        page_config = config_loader.get_page_settings()
        
        if not page_config:
            return True
            
        # 获取第一个分节来设置页面属性
        section = doc.sections[0]
        
        # 页边距设置
        margins = page_config.get("margins", {})
        if margins.get("top"):
            section.top_margin = config_loader.parse_length(margins["top"])
        if margins.get("bottom"):
            section.bottom_margin = config_loader.parse_length(margins["bottom"])
        if margins.get("left"):
            section.left_margin = config_loader.parse_length(margins["left"])
        if margins.get("right"):
            section.right_margin = config_loader.parse_length(margins["right"])
        
        # 页面方向
        orientation = page_config.get("orientation", "portrait")
        section.orientation = config_loader.parse_orientation(orientation)
        
        print("✅ 页面设置应用完成")
        return True
        
    except Exception as e:
        print(f"❌ 应用页面设置失败: {e}")
        return False