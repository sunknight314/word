from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import json

def extract_style_details(style):
    """提取样式的字体和段落格式信息（修复空值问题）"""
    style_info = {
        "name": style.name,
        "type": "段落" if style.type == WD_STYLE_TYPE.PARAGRAPH else "字符",
        "font": {},
        "paragraph": {}
    }

    # 提取字体样式（增加空值保护）
    font = style.font
    style_info["font"] = {
        "name": font.name,
        "east_asia_name": font._element.rPr.rFonts.get(qn("w:eastAsia"), None) 
            if font._element.rPr is not None and font._element.rPr.rFonts is not None else None,
        "size": f"{font.size.pt}磅" if font.size else None,
        "bold": font.bold,
        "italic": font.italic,
        "color": f"#{font.color.rgb}" if font.color and font.color.rgb else None
    }

    # 提取段落样式（仅段落样式有效）
    if style.type == WD_STYLE_TYPE.PARAGRAPH:
        para_format = style.paragraph_format
        style_info["paragraph"] = {
            "alignment": str(para_format.alignment).split('.')[-1] if para_format.alignment else None,
            "left_indent": f"{para_format.left_indent.pt}磅" if para_format.left_indent else None,
            "first_line_indent": f"{para_format.first_line_indent.pt}磅" if para_format.first_line_indent else None,
            "line_spacing": f"{para_format.line_spacing:.1f}倍" if para_format.line_spacing else None
        }
    return style_info

def extract_all_styles(docx_path):
    """提取文档所有预设样式"""
    doc = Document(docx_path)
    return [extract_style_details(style) for style in doc.styles if not style.hidden]

# 使用示例
if __name__ == "__main__":
    styles = extract_all_styles("../1、西安电子科技大学研究生学位论文模板（2015年版）-2022.11修订.docx")
    with open("样式信息.json", "w", encoding="utf-8") as f:
        json.dump(styles, f, ensure_ascii=False, indent=2)
    print(f"成功提取 {len(styles)} 种样式 → 样式信息.json")