"""
Word文档样式提取器
提取Word文档中的所有样式并输出详细的格式设置信息
"""

import json
from pathlib import Path
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.oxml.ns import qn


class WordStyleExtractor:
    """Word文档样式提取器"""
    
    def __init__(self):
        self.styles_info = {}
        
    def extract_styles_from_document(self, docx_path: str) -> dict:
        """
        从Word文档中提取所有样式信息
        
        Args:
            docx_path: Word文档路径
            
        Returns:
            包含所有样式信息的字典
        """
        try:
            doc = Document(docx_path)
            
            print(f"📖 正在提取文档样式: {Path(docx_path).name}")
            print(f"📊 文档中共有 {len(doc.styles)} 个样式")
            
            result = {
                "document_info": {
                    "file_name": Path(docx_path).name,
                    "total_styles": len(doc.styles)
                },
                "styles": {}
            }
            
            # 按样式类型分类统计
            style_type_count = {}
            
            for style in doc.styles:
                style_info = self._extract_style_info(style)
                result["styles"][style.name] = style_info
                
                # 统计样式类型
                style_type = style_info["type"]
                style_type_count[style_type] = style_type_count.get(style_type, 0) + 1
            
            result["document_info"]["style_type_count"] = style_type_count
            
            print("✅ 样式提取完成")
            return result
            
        except Exception as e:
            print(f"❌ 提取样式失败: {str(e)}")
            return {}
    
    def _extract_style_info(self, style) -> dict:
        """提取单个样式的详细信息"""
        style_info = {
            "name": style.name,
            "type": self._get_style_type_name(style.type),
            "builtin": style.builtin,
            "hidden": style.hidden,
            "locked": style.locked,
            "priority": style.priority,
            "unhide_when_used": style.unhide_when_used,
            "quick_style": style.quick_style
        }
        
        # 只处理段落样式和字符样式
        if style.type in [WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.CHARACTER]:
            # 提取字体信息
            if hasattr(style, 'font'):
                style_info["font"] = self._extract_font_info(style.font)
            
            # 提取段落格式信息（仅段落样式）
            if style.type == WD_STYLE_TYPE.PARAGRAPH and hasattr(style, 'paragraph_format'):
                style_info["paragraph"] = self._extract_paragraph_info(style.paragraph_format)
                
                # 提取大纲级别
                outline_level = self._get_outline_level(style)
                if outline_level is not None:
                    style_info["outline_level"] = outline_level
                
                # 提取样式继承信息
                style_info["base_style"] = self._get_base_style(style)
        
        return style_info
    
    def _get_style_type_name(self, style_type) -> str:
        """获取样式类型的名称"""
        type_map = {
            WD_STYLE_TYPE.PARAGRAPH: "段落样式",
            WD_STYLE_TYPE.CHARACTER: "字符样式", 
            WD_STYLE_TYPE.TABLE: "表格样式",
            WD_STYLE_TYPE.LIST: "列表样式"
        }
        return type_map.get(style_type, f"未知类型({style_type})")
    
    def _extract_font_info(self, font) -> dict:
        """提取详细的字体信息"""
        font_info = {}
        
        # 基本字体属性 - 即使为None也要记录，显示样式的完整状态
        font_info["name"] = font.name if font.name else None
        if font.size:
            font_info["size"] = self._safe_convert_to_pt(font.size)
            font_info["size_raw"] = font.size  # 保留原始值
        else:
            font_info["size"] = None
        if font.bold is not None:
            font_info["bold"] = font.bold
        if font.italic is not None:
            font_info["italic"] = font.italic
        if font.underline is not None:
            font_info["underline"] = str(font.underline)
            font_info["underline_type"] = font.underline
        if hasattr(font, 'strike') and font.strike is not None:
            font_info["strike"] = font.strike
        if hasattr(font, 'double_strike') and font.double_strike is not None:
            font_info["double_strike"] = font.double_strike
        if hasattr(font, 'outline') and font.outline is not None:
            font_info["outline"] = font.outline
        if hasattr(font, 'shadow') and font.shadow is not None:
            font_info["shadow"] = font.shadow
        if hasattr(font, 'emboss') and font.emboss is not None:
            font_info["emboss"] = font.emboss
        if hasattr(font, 'imprint') and font.imprint is not None:
            font_info["imprint"] = font.imprint
        if hasattr(font, 'small_caps') and font.small_caps is not None:
            font_info["small_caps"] = font.small_caps
        if hasattr(font, 'all_caps') and font.all_caps is not None:
            font_info["all_caps"] = font.all_caps
        if hasattr(font, 'hidden') and font.hidden is not None:
            font_info["hidden"] = font.hidden
        if hasattr(font, 'math') and font.math is not None:
            font_info["math"] = font.math
        if hasattr(font, 'no_proof') and font.no_proof is not None:
            font_info["no_proof"] = font.no_proof
        if hasattr(font, 'snap_to_grid') and font.snap_to_grid is not None:
            font_info["snap_to_grid"] = font.snap_to_grid
        if hasattr(font, 'spec_vanish') and font.spec_vanish is not None:
            font_info["spec_vanish"] = font.spec_vanish
        if hasattr(font, 'web_hidden') and font.web_hidden is not None:
            font_info["web_hidden"] = font.web_hidden
        
        # 颜色信息
        if font.color:
            if font.color.rgb:
                font_info["color_rgb"] = str(font.color.rgb)
            if font.color.theme_color:
                font_info["color_theme"] = str(font.color.theme_color)
            if font.color.type:
                font_info["color_type"] = str(font.color.type)
        
        if font.highlight_color is not None:
            font_info["highlight_color"] = str(font.highlight_color)
        
        # 字符间距（安全检查属性是否存在）
        if hasattr(font, 'kerning') and font.kerning is not None:
            font_info["kerning"] = self._safe_convert_to_pt(font.kerning)
        if hasattr(font, 'position') and font.position is not None:
            font_info["position"] = self._safe_convert_to_pt(font.position)
        if hasattr(font, 'scaling') and font.scaling is not None:
            font_info["scaling"] = font.scaling
        if hasattr(font, 'spacing') and font.spacing is not None:
            font_info["spacing"] = self._safe_convert_to_pt(font.spacing)
        
        # 提取中英文字体（通过XML）
        try:
            mixed_fonts = self._extract_mixed_font_info(font)
            if mixed_fonts:
                font_info.update(mixed_fonts)
        except:
            pass
        
        # 强制提取XML字体信息（即使基本属性为空）
        try:
            xml_fonts = self._force_extract_xml_fonts(font)
            if xml_fonts:
                font_info.update(xml_fonts)
        except:
            pass
        
        # 提取更多XML属性
        try:
            xml_attrs = self._extract_font_xml_attributes(font)
            if xml_attrs:
                font_info["xml_attributes"] = xml_attrs
        except:
            pass
            
        return font_info
    
    def _extract_mixed_font_info(self, font) -> dict:
        """提取中英文混合字体信息"""
        try:
            font_info = {}
            
            # 获取font的XML元素
            if hasattr(font, '_element'):
                font_element = font._element
                
                # 查找rFonts元素
                rfonts = font_element.find(qn('w:rFonts'))
                if rfonts is not None:
                    # 提取各种字体设置
                    ascii_font = rfonts.get(qn('w:ascii'))
                    if ascii_font:
                        font_info["ascii_font"] = ascii_font
                    
                    hansi_font = rfonts.get(qn('w:hAnsi'))
                    if hansi_font:
                        font_info["hansi_font"] = hansi_font
                    
                    eastasia_font = rfonts.get(qn('w:eastAsia'))
                    if eastasia_font:
                        font_info["eastasia_font"] = eastasia_font
                        
                    cs_font = rfonts.get(qn('w:cs'))
                    if cs_font:
                        font_info["cs_font"] = cs_font
            
            return font_info
        except:
            return {}
    
    def _extract_paragraph_info(self, para_format) -> dict:
        """提取详细的段落格式信息"""
        para_info = {}
        
        # 对齐方式
        if para_format.alignment is not None:
            para_info["alignment"] = str(para_format.alignment)
            para_info["alignment_value"] = para_format.alignment
        
        # 行距设置
        if para_format.line_spacing is not None:
            para_info["line_spacing"] = self._safe_convert_to_pt(para_format.line_spacing)
            para_info["line_spacing_raw"] = para_format.line_spacing
        if para_format.line_spacing_rule is not None:
            para_info["line_spacing_rule"] = str(para_format.line_spacing_rule)
            para_info["line_spacing_rule_value"] = para_format.line_spacing_rule
        
        # 段落间距
        if para_format.space_before is not None:
            para_info["space_before"] = self._safe_convert_to_pt(para_format.space_before)
            para_info["space_before_raw"] = para_format.space_before
        if para_format.space_after is not None:
            para_info["space_after"] = self._safe_convert_to_pt(para_format.space_after)
            para_info["space_after_raw"] = para_format.space_after
        if hasattr(para_format, 'space_before_auto') and para_format.space_before_auto is not None:
            para_info["space_before_auto"] = para_format.space_before_auto
        if hasattr(para_format, 'space_after_auto') and para_format.space_after_auto is not None:
            para_info["space_after_auto"] = para_format.space_after_auto
        
        # 缩进设置
        if para_format.first_line_indent is not None:
            para_info["first_line_indent"] = self._safe_convert_to_pt(para_format.first_line_indent)
            para_info["first_line_indent_raw"] = para_format.first_line_indent
        if para_format.left_indent is not None:
            para_info["left_indent"] = self._safe_convert_to_pt(para_format.left_indent)
            para_info["left_indent_raw"] = para_format.left_indent
        if para_format.right_indent is not None:
            para_info["right_indent"] = self._safe_convert_to_pt(para_format.right_indent)
            para_info["right_indent_raw"] = para_format.right_indent
        
        # 段落控制
        if para_format.widow_control is not None:
            para_info["widow_control"] = para_format.widow_control
        if para_format.keep_together is not None:
            para_info["keep_together"] = para_format.keep_together
        if para_format.keep_with_next is not None:
            para_info["keep_with_next"] = para_format.keep_with_next
        if para_format.page_break_before is not None:
            para_info["page_break_before"] = para_format.page_break_before
        
        # 制表符设置
        try:
            if para_format.tab_stops:
                tab_stops_info = []
                for tab_stop in para_format.tab_stops:
                    tab_info = {
                        "position": self._safe_convert_to_pt(tab_stop.position),
                        "alignment": str(tab_stop.alignment) if tab_stop.alignment else None,
                        "leader": str(tab_stop.leader) if tab_stop.leader else None
                    }
                    tab_stops_info.append(tab_info)
                para_info["tab_stops"] = tab_stops_info
        except:
            pass
        
        # 边框设置（通过XML提取）
        try:
            border_info = self._extract_paragraph_borders(para_format)
            if border_info:
                para_info["borders"] = border_info
        except:
            pass
        
        # 阴影设置
        try:
            shading_info = self._extract_paragraph_shading(para_format)
            if shading_info:
                para_info["shading"] = shading_info
        except:
            pass
        
        # 提取更多XML属性
        try:
            xml_attrs = self._extract_paragraph_xml_attributes(para_format)
            if xml_attrs:
                para_info["xml_attributes"] = xml_attrs
        except:
            pass
        
        # 强制提取段落XML信息
        try:
            xml_para_info = self._force_extract_paragraph_xml(para_format)
            if xml_para_info:
                para_info["xml_detailed"] = xml_para_info
        except:
            pass
        
        return para_info
    
    def _get_outline_level(self, style) -> int:
        """获取样式的大纲级别"""
        try:
            if hasattr(style, '_element'):
                style_element = style._element
                ppr = style_element.find(qn('w:pPr'))
                if ppr is not None:
                    outline_lvl = ppr.find(qn('w:outlineLvl'))
                    if outline_lvl is not None:
                        return int(outline_lvl.get(qn('w:val'), 0))
            return None
        except:
            return None
    
    def _safe_convert_to_pt(self, value) -> str:
        """安全地将尺寸值转换为pt字符串"""
        try:
            if value is None:
                return "0pt"
            
            # 如果已经是Length对象，直接获取pt值
            if hasattr(value, 'pt'):
                return f"{value.pt}pt"
            
            # 如果是数值，假设单位是twips（1pt = 20 twips）
            if isinstance(value, (int, float)):
                pt_value = value / 20.0  # 将twips转换为pt
                return f"{pt_value}pt"
            
            # 其他情况，转换为字符串
            return str(value)
        except:
            return "0pt"
    
    def _safe_get_attribute(self, obj, attr_name, default=None):
        """安全地获取对象属性"""
        try:
            if hasattr(obj, attr_name):
                value = getattr(obj, attr_name)
                return value if value is not None else default
            return default
        except:
            return default
    
    def _get_base_style(self, style):
        """获取样式的基础样式（继承关系）"""
        try:
            if hasattr(style, 'base_style') and style.base_style:
                return {
                    "name": style.base_style.name,
                    "type": self._get_style_type_name(style.base_style.type)
                }
            return None
        except:
            return None
    
    def _force_extract_xml_fonts(self, font) -> dict:
        """强制从XML中提取字体信息，即使基本属性为空"""
        xml_fonts = {}
        try:
            if hasattr(font, '_element'):
                font_element = font._element
                
                # 查找rFonts元素
                rfonts = font_element.find(qn('w:rFonts'))
                if rfonts is not None:
                    # 强制提取所有字体族，即使为空也记录
                    xml_fonts["ascii_font"] = rfonts.get(qn('w:ascii'))
                    xml_fonts["hansi_font"] = rfonts.get(qn('w:hAnsi'))
                    xml_fonts["eastasia_font"] = rfonts.get(qn('w:eastAsia'))
                    xml_fonts["cs_font"] = rfonts.get(qn('w:cs'))
                    
                    # 添加字体族解释
                    font_descriptions = {}
                    if xml_fonts["ascii_font"]:
                        font_descriptions["ascii_description"] = f"ASCII字符字体: {xml_fonts['ascii_font']}"
                    if xml_fonts["hansi_font"]:
                        font_descriptions["hansi_description"] = f"高ANSI字符字体: {xml_fonts['hansi_font']}"
                    if xml_fonts["eastasia_font"]:
                        font_descriptions["eastasia_description"] = f"东亚字符字体(中日韩): {xml_fonts['eastasia_font']}"
                    if xml_fonts["cs_font"]:
                        font_descriptions["cs_description"] = f"复杂脚本字体: {xml_fonts['cs_font']}"
                    
                    xml_fonts.update(font_descriptions)
                
                # 检查其他字体属性
                sz = font_element.find(qn('w:sz'))
                if sz is not None:
                    xml_fonts["size_half_points"] = sz.get(qn('w:val'))
                    xml_fonts["size_points"] = float(sz.get(qn('w:val'))) / 2.0
                
                # 检查字体效果
                if font_element.find(qn('w:b')) is not None:
                    xml_fonts["bold_xml"] = True
                if font_element.find(qn('w:i')) is not None:
                    xml_fonts["italic_xml"] = True
                if font_element.find(qn('w:u')) is not None:
                    u_elem = font_element.find(qn('w:u'))
                    xml_fonts["underline_xml"] = u_elem.get(qn('w:val'), 'single')
                
                # 检查颜色
                color_elem = font_element.find(qn('w:color'))
                if color_elem is not None:
                    xml_fonts["color_xml"] = color_elem.get(qn('w:val'))
                    xml_fonts["theme_color_xml"] = color_elem.get(qn('w:themeColor'))
                
        except Exception as e:
            xml_fonts["xml_extraction_error"] = str(e)
        
        return xml_fonts
    
    def _extract_font_xml_attributes(self, font) -> dict:
        """提取字体的XML属性"""
        xml_attrs = {}
        try:
            if hasattr(font, '_element'):
                font_element = font._element
                
                # 提取字体缩放
                sz = font_element.find(qn('w:sz'))
                if sz is not None:
                    xml_attrs["font_size_half_points"] = sz.get(qn('w:val'))
                
                # 提取字体变形
                w_elem = font_element.find(qn('w:w'))
                if w_elem is not None:
                    xml_attrs["font_scale"] = w_elem.get(qn('w:val'))
                
                # 提取字符间距
                spacing_elem = font_element.find(qn('w:spacing'))
                if spacing_elem is not None:
                    xml_attrs["character_spacing"] = spacing_elem.get(qn('w:val'))
                
                # 提取位置偏移
                position_elem = font_element.find(qn('w:position'))
                if position_elem is not None:
                    xml_attrs["vertical_position"] = position_elem.get(qn('w:val'))
        except:
            pass
        return xml_attrs
    
    def _extract_paragraph_borders(self, para_format) -> dict:
        """提取段落边框信息"""
        border_info = {}
        try:
            if hasattr(para_format, '_element'):
                para_element = para_format._element
                ppr = para_element.find(qn('w:pPr'))
                if ppr is not None:
                    pBdr = ppr.find(qn('w:pBdr'))
                    if pBdr is not None:
                        # 提取各边框信息
                        borders = ["top", "left", "bottom", "right", "between", "bar"]
                        for border in borders:
                            border_elem = pBdr.find(qn(f'w:{border}'))
                            if border_elem is not None:
                                border_info[border] = {
                                    "style": border_elem.get(qn('w:val')),
                                    "size": border_elem.get(qn('w:sz')),
                                    "space": border_elem.get(qn('w:space')),
                                    "color": border_elem.get(qn('w:color'))
                                }
        except:
            pass
        return border_info
    
    def _extract_paragraph_shading(self, para_format) -> dict:
        """提取段落阴影信息"""
        shading_info = {}
        try:
            if hasattr(para_format, '_element'):
                para_element = para_format._element
                ppr = para_element.find(qn('w:pPr'))
                if ppr is not None:
                    shd = ppr.find(qn('w:shd'))
                    if shd is not None:
                        shading_info = {
                            "pattern": shd.get(qn('w:val')),
                            "color": shd.get(qn('w:color')),
                            "fill": shd.get(qn('w:fill')),
                            "theme_color": shd.get(qn('w:themeColor')),
                            "theme_fill": shd.get(qn('w:themeFill'))
                        }
        except:
            pass
        return shading_info
    
    def _extract_paragraph_xml_attributes(self, para_format) -> dict:
        """提取段落的XML属性"""
        xml_attrs = {}
        try:
            if hasattr(para_format, '_element'):
                para_element = para_format._element
                ppr = para_element.find(qn('w:pPr'))
                if ppr is not None:
                    # 提取编号信息
                    numPr = ppr.find(qn('w:numPr'))
                    if numPr is not None:
                        ilvl = numPr.find(qn('w:ilvl'))
                        numId = numPr.find(qn('w:numId'))
                        xml_attrs["numbering"] = {
                            "level": ilvl.get(qn('w:val')) if ilvl is not None else None,
                            "id": numId.get(qn('w:val')) if numId is not None else None
                        }
                    
                    # 提取样式链接
                    pStyle = ppr.find(qn('w:pStyle'))
                    if pStyle is not None:
                        xml_attrs["paragraph_style_id"] = pStyle.get(qn('w:val'))
                    
                    # 提取文本方向
                    textDirection = ppr.find(qn('w:textDirection'))
                    if textDirection is not None:
                        xml_attrs["text_direction"] = textDirection.get(qn('w:val'))
                    
                    # 提取双向文本
                    bidi = ppr.find(qn('w:bidi'))
                    if bidi is not None:
                        xml_attrs["bidirectional"] = bidi.get(qn('w:val'), 'true')
        except:
            pass
        return xml_attrs
    
    def _force_extract_paragraph_xml(self, para_format) -> dict:
        """强制从XML中提取段落信息，即使基本属性为空"""
        xml_para = {}
        try:
            if hasattr(para_format, '_element'):
                para_element = para_format._element
                ppr = para_element.find(qn('w:pPr'))
                if ppr is not None:
                    # 强制提取对齐方式
                    jc = ppr.find(qn('w:jc'))
                    if jc is not None:
                        xml_para["alignment_xml"] = jc.get(qn('w:val'))
                    
                    # 强制提取间距设置
                    spacing = ppr.find(qn('w:spacing'))
                    if spacing is not None:
                        xml_para["spacing_xml"] = {
                            "before": spacing.get(qn('w:before')),
                            "after": spacing.get(qn('w:after')),
                            "line": spacing.get(qn('w:line')),
                            "line_rule": spacing.get(qn('w:lineRule'))
                        }
                    
                    # 强制提取缩进设置
                    ind = ppr.find(qn('w:ind'))
                    if ind is not None:
                        xml_para["indentation_xml"] = {
                            "left": ind.get(qn('w:left')),
                            "right": ind.get(qn('w:right')),
                            "first_line": ind.get(qn('w:firstLine')),
                            "hanging": ind.get(qn('w:hanging'))
                        }
                    
                    # 强制提取边框设置
                    pBdr = ppr.find(qn('w:pBdr'))
                    if pBdr is not None:
                        borders_xml = {}
                        for border_type in ["top", "left", "bottom", "right", "between", "bar"]:
                            border_elem = pBdr.find(qn(f'w:{border_type}'))
                            if border_elem is not None:
                                borders_xml[border_type] = {
                                    "val": border_elem.get(qn('w:val')),
                                    "sz": border_elem.get(qn('w:sz')),
                                    "space": border_elem.get(qn('w:space')),
                                    "color": border_elem.get(qn('w:color'))
                                }
                        if borders_xml:
                            xml_para["borders_xml"] = borders_xml
                    
                    # 强制提取阴影设置
                    shd = ppr.find(qn('w:shd'))
                    if shd is not None:
                        xml_para["shading_xml"] = {
                            "val": shd.get(qn('w:val')),
                            "color": shd.get(qn('w:color')),
                            "fill": shd.get(qn('w:fill')),
                            "theme_color": shd.get(qn('w:themeColor')),
                            "theme_fill": shd.get(qn('w:themeFill'))
                        }
                    
                    # 强制提取制表符设置
                    tabs = ppr.find(qn('w:tabs'))
                    if tabs is not None:
                        tabs_xml = []
                        for tab in tabs.findall(qn('w:tab')):
                            tabs_xml.append({
                                "val": tab.get(qn('w:val')),
                                "pos": tab.get(qn('w:pos')),
                                "leader": tab.get(qn('w:leader'))
                            })
                        if tabs_xml:
                            xml_para["tabs_xml"] = tabs_xml
                    
                    # 强制提取段落控制
                    if ppr.find(qn('w:widowControl')) is not None:
                        xml_para["widow_control_xml"] = ppr.find(qn('w:widowControl')).get(qn('w:val'), 'true')
                    if ppr.find(qn('w:keepNext')) is not None:
                        xml_para["keep_next_xml"] = ppr.find(qn('w:keepNext')).get(qn('w:val'), 'true')
                    if ppr.find(qn('w:keepLines')) is not None:
                        xml_para["keep_lines_xml"] = ppr.find(qn('w:keepLines')).get(qn('w:val'), 'true')
                    if ppr.find(qn('w:pageBreakBefore')) is not None:
                        xml_para["page_break_before_xml"] = ppr.find(qn('w:pageBreakBefore')).get(qn('w:val'), 'true')
                    
                    # 强制提取大纲级别
                    outlineLvl = ppr.find(qn('w:outlineLvl'))
                    if outlineLvl is not None:
                        xml_para["outline_level_xml"] = outlineLvl.get(qn('w:val'))
        
        except Exception as e:
            xml_para["xml_extraction_error"] = str(e)
        
        return xml_para
    
    def save_to_json(self, styles_data: dict, output_path: str):
        """将样式信息保存为JSON文件"""
        try:
            # 清理数据，确保所有对象都能序列化
            clean_data = self._clean_data_for_json(styles_data)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(clean_data, f, ensure_ascii=False, indent=2, default=str)
            
            file_size = Path(output_path).stat().st_size
            print(f"💾 样式信息已保存到: {output_path} ({file_size} 字节)")
        except Exception as e:
            print(f"❌ 保存JSON文件失败: {str(e)}")
    
    def _clean_data_for_json(self, data):
        """清理数据使其能够被JSON序列化"""
        if isinstance(data, dict):
            return {key: self._clean_data_for_json(value) for key, value in data.items()}
        elif isinstance(data, list):
            return [self._clean_data_for_json(item) for item in data]
        elif hasattr(data, '__dict__'):
            # 对于有属性的对象，转换为字符串
            return str(data)
        else:
            return data
    
    def print_summary(self, styles_data: dict):
        """打印样式摘要信息"""
        if not styles_data:
            return
            
        print("\n" + "="*60)
        print("📋 样式提取摘要")
        print("="*60)
        
        doc_info = styles_data.get("document_info", {})
        print(f"📄 文档名称: {doc_info.get('file_name', 'Unknown')}")
        print(f"📊 样式总数: {doc_info.get('total_styles', 0)}")
        
        # 打印样式类型统计
        type_count = doc_info.get("style_type_count", {})
        if type_count:
            print("\n📈 样式类型统计:")
            for style_type, count in type_count.items():
                print(f"  {style_type}: {count}个")
        
        # 列出详细样式信息
        styles = styles_data.get("styles", {})
        if styles:
            print(f"\n📝 详细样式信息 (前10个):")
            for i, (name, info) in enumerate(list(styles.items())[:10]):
                builtin = "✅" if info.get("builtin", False) else "❌"
                style_type = info.get('type', 'Unknown')
                
                # 提取关键信息
                details = []
                if 'font' in info:
                    font_info = info['font']
                    if 'size' in font_info:
                        details.append(f"字号:{font_info['size']}")
                    if 'name' in font_info:
                        details.append(f"字体:{font_info['name']}")
                    if font_info.get('bold'):
                        details.append("粗体")
                    if font_info.get('italic'):
                        details.append("斜体")
                
                if 'paragraph' in info:
                    para_info = info['paragraph']
                    if 'alignment' in para_info:
                        details.append(f"对齐:{para_info['alignment']}")
                    if 'first_line_indent' in para_info and para_info['first_line_indent'] != "0pt":
                        details.append(f"首行缩进:{para_info['first_line_indent']}")
                
                if info.get('outline_level') is not None:
                    details.append(f"大纲级别:{info['outline_level']}")
                
                detail_str = ", ".join(details[:3])  # 只显示前3个细节
                if len(details) > 3:
                    detail_str += "..."
                
                print(f"  {i+1:2d}. {name}")
                print(f"      类型:{style_type} | 内置:{builtin}")
                if detail_str:
                    print(f"      详情:{detail_str}")
            
            if len(styles) > 10:
                print(f"\n  ... 还有 {len(styles) - 10} 个样式")
            
            # 统计有详细设置的样式
            detailed_count = 0
            for info in styles.values():
                if 'font' in info or 'paragraph' in info:
                    detailed_count += 1
            
            print(f"\n📈 统计信息:")
            print(f"  包含详细格式设置的样式: {detailed_count}个")
            print(f"  包含字体设置的样式: {sum(1 for info in styles.values() if 'font' in info)}个")
            print(f"  包含段落设置的样式: {sum(1 for info in styles.values() if 'paragraph' in info)}个")
            print(f"  包含大纲级别的样式: {sum(1 for info in styles.values() if info.get('outline_level') is not None)}个")


def main():
    """主函数 - 演示样式提取功能"""
    extractor = WordStyleExtractor()
    
    # 测试文件路径 - 请修改为你的Word文档路径
    test_files = [
        "../test.docx"
    ]
    
    for doc_path in test_files:
        if Path(doc_path).exists():
            print(f"\n🔍 正在处理: {doc_path}")
            
            # 提取样式
            styles_data = extractor.extract_styles_from_document(doc_path)
            
            if styles_data:
                # 打印摘要
                extractor.print_summary(styles_data)
                
                # 保存JSON文件
                doc_name = Path(doc_path).stem
                output_file = f"styles_{doc_name}.json"
                extractor.save_to_json(styles_data, output_file)
                
                print(f"\n✅ {doc_path} 处理完成")
            else:
                print(f"❌ {doc_path} 处理失败")
        else:
            print(f"⚠️  文件不存在: {doc_path}")


if __name__ == "__main__":
    main()