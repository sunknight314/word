"""
Word文档样式提取器
使用python-docx库直接提取Word文档中的样式信息
"""

import json
from typing import Dict, List, Any, Optional
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

from ..models.format_schema import (
    DocumentFormat, PageSetup, Margins, ParagraphStyle, 
    FontStyle, CharacterStyle, TableStyle, ListStyle
)

class StyleExtractor:
    """Word文档样式提取器"""
    
    def __init__(self):
        self.alignment_map = {
            WD_ALIGN_PARAGRAPH.LEFT: "left",
            WD_ALIGN_PARAGRAPH.CENTER: "center", 
            WD_ALIGN_PARAGRAPH.RIGHT: "right",
            WD_ALIGN_PARAGRAPH.JUSTIFY: "justify"
        }
    
    def extract_styles_from_document(self, file_path: str) -> Dict[str, Any]:
        """
        从Word文档中提取所有样式信息
        
        Args:
            file_path: Word文档路径
            
        Returns:
            格式化的样式信息字典
        """
        try:
            doc = Document(file_path)
            
            # 提取页面设置
            page_setup = self._extract_page_setup(doc)
            
            # 提取段落样式
            paragraph_styles = self._extract_paragraph_styles(doc)
            
            # 提取字符样式
            character_styles = self._extract_character_styles(doc)
            
            # 提取表格样式
            table_styles = self._extract_table_styles(doc)
            
            # 分析文档内容获取实际使用的样式
            content_analysis = self._analyze_document_content(doc)
            
            # 构建完整格式信息
            format_info = {
                "document_name": file_path.split("/")[-1],
                "extraction_method": "python-docx",
                "page_setup": page_setup,
                "paragraph_styles": paragraph_styles,
                "character_styles": character_styles,
                "table_styles": table_styles,
                "content_analysis": content_analysis,
                "style_summary": self._generate_style_summary(paragraph_styles, character_styles)
            }
            
            return format_info
            
        except Exception as e:
            raise Exception(f"样式提取失败: {str(e)}")
    
    def _extract_page_setup(self, doc: Document) -> Dict[str, Any]:
        """提取页面设置信息"""
        section = doc.sections[0]
        print("!!!!!!!!!!!!!!!!!!!!")
        print(section.top_margin.cm)
        
        return {
            "margins": {
                "top": round(section.top_margin.cm, 2),
                "bottom": round(section.bottom_margin.cm, 2),
                "left": round(section.left_margin.cm, 2),
                "right": round(section.right_margin.cm, 2)
            },
            "orientation": "landscape" if section.orientation == 1 else "portrait",
            "page_width": round(section.page_width.cm, 2),
            "page_height": round(section.page_height.cm, 2),
            "paper_size": self._detect_paper_size(section.page_width.cm, section.page_height.cm)
        }
    
    def _extract_paragraph_styles(self, doc: Document) -> Dict[str, Any]:
        """提取段落样式"""
        paragraph_styles = {}
        
        for style in doc.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH:
                style_info = self._parse_paragraph_style(style)
                if style_info:
                    paragraph_styles[style.name] = style_info
        
        return paragraph_styles
    
    def _extract_character_styles(self, doc: Document) -> Dict[str, Any]:
        """提取字符样式"""
        character_styles = {}
        
        for style in doc.styles:
            if style.type == WD_STYLE_TYPE.CHARACTER:
                style_info = self._parse_character_style(style)
                if style_info:
                    character_styles[style.name] = style_info
        
        return character_styles
    
    def _extract_table_styles(self, doc: Document) -> Dict[str, Any]:
        """提取表格样式"""
        table_styles = {}
        
        for style in doc.styles:
            if style.type == WD_STYLE_TYPE.TABLE:
                style_info = self._parse_table_style(style)
                if style_info:
                    table_styles[style.name] = style_info
        
        return table_styles
    
    def _parse_paragraph_style(self, style) -> Optional[Dict[str, Any]]:
        """解析段落样式"""
        try:
            paragraph_format = style.paragraph_format
            font = style.font
            print("style name:" + style.name)
            print("!!!!!!name")
            print(font.name)
            
            # 提取字体信息
            font_info = {
                "name": font.name or "宋体",
                "size": round(font.size.pt, 1) if font.size else 12,
                "bold": font.bold or False,
                "italic": font.italic or False,
                "underline": font.underline or False,
                "color": self._get_font_color(font)
            }
            
            # 提取段落格式
            style_info = {
                "name": style.name,
                "font": font_info,
                "alignment": self.alignment_map.get(paragraph_format.alignment, "left"),
                "line_spacing": self._get_line_spacing(paragraph_format),
                "space_before": round(paragraph_format.space_before.pt, 1) if paragraph_format.space_before else 0,
                "space_after": round(paragraph_format.space_after.pt, 1) if paragraph_format.space_after else 0,
                "left_indent": round(paragraph_format.left_indent.cm, 2) if paragraph_format.left_indent else 0,
                "right_indent": round(paragraph_format.right_indent.cm, 2) if paragraph_format.right_indent else 0,
                "first_line_indent": round(paragraph_format.first_line_indent.cm, 2) if paragraph_format.first_line_indent else 0,
                "keep_with_next": paragraph_format.keep_with_next or False,
                "keep_together": paragraph_format.keep_together or False
            }
            
            return style_info
            
        except Exception as e:
            print(f"解析段落样式 {style.name} 失败: {e}")
            return None
    
    def _parse_character_style(self, style) -> Optional[Dict[str, Any]]:
        """解析字符样式"""
        try:
            font = style.font
            
            font_info = {
                "name": font.name or "宋体",
                "size": round(font.size.pt, 1) if font.size else 12,
                "bold": font.bold or False,
                "italic": font.italic or False,
                "underline": font.underline or False,
                "color": self._get_font_color(font)
            }
            
            return {
                "name": style.name,
                "font": font_info,
                "highlight_color": self._get_highlight_color(font)
            }
            
        except Exception as e:
            print(f"解析字符样式 {style.name} 失败: {e}")
            return None
    
    def _parse_table_style(self, style) -> Optional[Dict[str, Any]]:
        """解析表格样式"""
        try:
            return {
                "name": style.name,
                "border_style": "single",  # 默认值，需要进一步解析
                "border_width": 0.5,
                "border_color": "#000000"
            }
        except Exception as e:
            print(f"解析表格样式 {style.name} 失败: {e}")
            return None
    
    def _analyze_document_content(self, doc: Document) -> Dict[str, Any]:
        """分析文档内容，统计样式使用情况"""
        style_usage = {}
        total_paragraphs = 0
        
        for paragraph in doc.paragraphs:
            total_paragraphs += 1
            style_name = paragraph.style.name
            style_usage[style_name] = style_usage.get(style_name, 0) + 1
        
        return {
            "total_paragraphs": total_paragraphs,
            "style_usage": style_usage,
            "most_used_styles": sorted(style_usage.items(), key=lambda x: x[1], reverse=True)[:5]
        }
    
    def _generate_style_summary(self, paragraph_styles: Dict, character_styles: Dict) -> Dict[str, Any]:
        """生成样式摘要"""
        return {
            "total_paragraph_styles": len(paragraph_styles),
            "total_character_styles": len(character_styles),
            "main_font_families": self._get_main_fonts(paragraph_styles),
            "font_size_range": self._get_font_size_range(paragraph_styles)
        }
    
    def _get_font_color(self, font) -> Optional[str]:
        """获取字体颜色"""
        try:
            if font.color and font.color.rgb:
                return f"#{font.color.rgb}"
            return None
        except:
            return None
    
    def _get_highlight_color(self, font) -> Optional[str]:
        """获取高亮颜色"""
        try:
            if hasattr(font, 'highlight_color') and font.highlight_color:
                return str(font.highlight_color)
            return None
        except:
            return None
    
    def _get_line_spacing(self, paragraph_format) -> float:
        """获取行间距"""
        try:
            if paragraph_format.line_spacing:
                if paragraph_format.line_spacing_rule == 1:  # 单倍行距
                    return paragraph_format.line_spacing
                else:
                    return 1.0
            return 1.0
        except:
            return 1.0
    
    def _detect_paper_size(self, width: float, height: float) -> str:
        """检测纸张大小"""
        # A4: 21.0 x 29.7 cm
        if abs(width - 21.0) < 0.5 and abs(height - 29.7) < 0.5:
            return "A4"
        # A3: 29.7 x 42.0 cm  
        elif abs(width - 29.7) < 0.5 and abs(height - 42.0) < 0.5:
            return "A3"
        # Letter: 21.6 x 27.9 cm
        elif abs(width - 21.6) < 0.5 and abs(height - 27.9) < 0.5:
            return "Letter"
        else:
            return "Custom"
    
    def _get_main_fonts(self, paragraph_styles: Dict) -> List[str]:
        """获取主要字体"""
        fonts = set()
        for style in paragraph_styles.values():
            if isinstance(style, dict) and 'font' in style:
                fonts.add(style['font']['name'])
        return list(fonts)
    
    def _get_font_size_range(self, paragraph_styles: Dict) -> Dict[str, float]:
        """获取字体大小范围"""
        sizes = []
        for style in paragraph_styles.values():
            if isinstance(style, dict) and 'font' in style:
                sizes.append(style['font']['size'])
        
        if sizes:
            return {
                "min_size": min(sizes),
                "max_size": max(sizes),
                "avg_size": round(sum(sizes) / len(sizes), 1)
            }
        return {"min_size": 12, "max_size": 12, "avg_size": 12}