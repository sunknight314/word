"""
格式展示助手 - 将格式配置转换为易于前端展示的格式
"""

from typing import Dict, Any, List


class FormatDisplayHelper:
    """格式展示助手类"""
    
    def format_for_display(self, format_config: Dict[str, Any]) -> Dict[str, Any]:
        """
        将格式配置转换为易于前端展示的格式
        
        Args:
            format_config: 原始格式配置
            
        Returns:
            格式化后的展示数据
        """
        display_data = {
            "page_settings": self._format_page_settings(format_config.get("page_settings", {})),
            "styles": self._format_styles(format_config.get("styles", {})),
            "document_structure": self._format_document_structure(format_config)
        }
        
        return display_data
    
    def _format_page_settings(self, page_settings: Dict[str, Any]) -> Dict[str, Any]:
        """格式化页面设置"""
        margins = page_settings.get("margins", {})
        
        return {
            "title": "页面设置",
            "items": [
                {
                    "label": "页面尺寸",
                    "value": page_settings.get("size", "A4"),
                    "category": "basic"
                },
                {
                    "label": "页面方向",
                    "value": "纵向" if page_settings.get("orientation") == "portrait" else "横向",
                    "category": "basic"
                },
                {
                    "label": "上边距",
                    "value": self._format_value(margins.get("top", "0pt")),
                    "category": "margins"
                },
                {
                    "label": "下边距",
                    "value": self._format_value(margins.get("bottom", "0pt")),
                    "category": "margins"
                },
                {
                    "label": "左边距",
                    "value": self._format_value(margins.get("left", "0pt")),
                    "category": "margins"
                },
                {
                    "label": "右边距",
                    "value": self._format_value(margins.get("right", "0pt")),
                    "category": "margins"
                }
            ]
        }
    
    def _format_styles(self, styles: Dict[str, Any]) -> List[Dict[str, Any]]:
        """格式化样式设置"""
        style_names = {
            "title": "文档标题",
            "heading1": "一级标题",
            "heading2": "二级标题",
            "heading3": "三级标题",
            "heading4": "四级标题",
            "paragraph": "正文段落",
            "abstract_title_cn": "中文摘要标题",
            "abstract_title_en": "英文摘要标题",
            "figure_caption": "图注",
            "table_caption": "表注"
        }
        
        formatted_styles = []
        
        for style_key, style_data in styles.items():
            if isinstance(style_data, dict):
                style_display = {
                    "name": style_names.get(style_key, style_key),
                    "key": style_key,
                    "font": self._format_font_settings(style_data.get("font", {})),
                    "paragraph": self._format_paragraph_settings(style_data.get("paragraph", {}))
                }
                formatted_styles.append(style_display)
        
        return formatted_styles
    
    def _format_font_settings(self, font: Dict[str, Any]) -> List[Dict[str, str]]:
        """格式化字体设置"""
        settings = []
        
        if "chinese" in font:
            settings.append({
                "label": "中文字体",
                "value": font["chinese"]
            })
        
        if "english" in font:
            settings.append({
                "label": "英文字体",
                "value": font["english"]
            })
        
        if "size" in font:
            settings.append({
                "label": "字号",
                "value": self._format_value(font["size"])
            })
        
        if "bold" in font:
            settings.append({
                "label": "加粗",
                "value": "是" if font["bold"] else "否"
            })
        
        if "italic" in font:
            settings.append({
                "label": "斜体",
                "value": "是" if font["italic"] else "否"
            })
        
        return settings
    
    def _format_paragraph_settings(self, paragraph: Dict[str, Any]) -> List[Dict[str, str]]:
        """格式化段落设置"""
        settings = []
        
        alignment_map = {
            "left": "左对齐",
            "center": "居中",
            "right": "右对齐",
            "justify": "两端对齐"
        }
        
        if "alignment" in paragraph:
            settings.append({
                "label": "对齐方式",
                "value": alignment_map.get(paragraph["alignment"], paragraph["alignment"])
            })
        
        if "line_spacing" in paragraph:
            settings.append({
                "label": "行距",
                "value": self._format_value(paragraph["line_spacing"])
            })
        
        if "space_before" in paragraph and paragraph["space_before"] != "0pt":
            settings.append({
                "label": "段前距",
                "value": self._format_value(paragraph["space_before"])
            })
        
        if "space_after" in paragraph and paragraph["space_after"] != "0pt":
            settings.append({
                "label": "段后距",
                "value": self._format_value(paragraph["space_after"])
            })
        
        if "first_line_indent" in paragraph and paragraph["first_line_indent"] != "0pt":
            settings.append({
                "label": "首行缩进",
                "value": self._format_value(paragraph["first_line_indent"])
            })
        
        if "hanging_indent" in paragraph and paragraph["hanging_indent"] != "0pt":
            settings.append({
                "label": "悬挂缩进",
                "value": self._format_value(paragraph["hanging_indent"])
            })
        
        return settings
    
    def _format_document_structure(self, format_config: Dict[str, Any]) -> Dict[str, Any]:
        """格式化文档结构设置"""
        items = []
        
        # TOC设置
        if "toc_settings" in format_config:
            toc = format_config["toc_settings"]
            if "title" in toc:
                items.append({
                    "label": "目录标题",
                    "value": toc["title"],
                    "category": "toc"
                })
            if "levels" in toc:
                items.append({
                    "label": "目录层级",
                    "value": toc["levels"],
                    "category": "toc"
                })
        
        # 页码设置
        if "page_numbering" in format_config:
            numbering = format_config["page_numbering"]
            if "toc_section" in numbering and "format" in numbering["toc_section"]:
                format_map = {
                    "upperRoman": "大写罗马数字",
                    "lowerRoman": "小写罗马数字",
                    "decimal": "阿拉伯数字"
                }
                items.append({
                    "label": "目录页码格式",
                    "value": format_map.get(numbering["toc_section"]["format"], 
                                          numbering["toc_section"]["format"]),
                    "category": "numbering"
                })
        
        return {
            "title": "文档结构",
            "items": items
        }
    
    def _format_value(self, value: str) -> str:
        """格式化数值，转换单位为易读格式"""
        if not value:
            return "未设置"
        
        # 转换pt为更易读的格式
        if isinstance(value, str) and value.endswith("pt"):
            try:
                pt_value = float(value[:-2])
                
                # 常见pt值的转换
                if pt_value == 0:
                    return "0"
                elif pt_value == 12:
                    return "12磅（小四）"
                elif pt_value == 14:
                    return "14磅（四号）"
                elif pt_value == 15:
                    return "15磅（小三）"
                elif pt_value == 16:
                    return "16磅（三号）"
                elif pt_value == 18:
                    return "18磅（小二）"
                elif pt_value == 20:
                    return "20磅（固定值）"
                elif pt_value == 22:
                    return "22磅（二号）"
                elif pt_value == 24:
                    return "24磅（小一/2字符）"
                elif pt_value == 26:
                    return "26磅（一号）"
                elif pt_value == 28.35:  # 约1cm
                    return "1厘米"
                elif pt_value == 56.7:   # 约2cm
                    return "2厘米"
                elif pt_value == 72:     # 1英寸
                    return "1英寸（2.54厘米）"
                else:
                    # 尝试转换为厘米
                    cm_value = pt_value / 28.35
                    if cm_value == round(cm_value, 1):
                        return f"{cm_value:.1f}厘米"
                    else:
                        return f"{pt_value:.1f}磅"
            except:
                pass
        
        return value