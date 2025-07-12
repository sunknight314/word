"""
文档格式转换服务 V2
按步骤进行完整的格式转换：
1. 页面设置
2. 创建样式
3. 应用样式
4. 创建章节（分节、页码、页眉页脚）
5. 创建目录
6. 保存文档
"""
import os
from typing import Dict, Any, List, Optional
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from app.services.analysis_result_storage import AnalysisResultStorage
from app.utils.unit_converter import UnitConverter
from app.constants.style_names import StyleNames, get_style_description
import logging

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class DocumentFormatterV2:
    """文档格式化类 V2"""
    
    def __init__(self):
        self.analysis_storage = AnalysisResultStorage()
        self.unit_converter = UnitConverter()
        self.doc = None
        self.format_config = None
        self.analysis_result = None
        
    def format_document(self, source_file_path: str, source_file_id: str, format_file_id: str) -> Dict[str, Any]:
        """
        格式化文档主流程
        """
        try:
            # 准备工作：加载数据
            logger.info("开始文档格式转换...")
            preparation_result = self._prepare_formatting(source_file_path, source_file_id, format_file_id)
            if not preparation_result["success"]:
                return preparation_result
            
            # 打开文档
            self.doc = Document(source_file_path)
            
            # 步骤记录
            steps_results = {}
            
            # 步骤1：应用页面设置
            logger.info("步骤1：应用页面设置")
            page_result = self._step1_apply_page_settings()
            steps_results["page_settings"] = page_result
            
            # 步骤2：创建样式
            logger.info("步骤2：创建样式")
            styles_result = self._step2_create_styles()
            steps_results["create_styles"] = styles_result
            
            # 步骤3：应用样式到段落
            logger.info("步骤3：应用样式")
            apply_result = self._step3_apply_styles()
            steps_results["apply_styles"] = apply_result
            
            # 步骤4：创建章节（分节、页码、页眉页脚）
            logger.info("步骤4：创建章节")
            section_result = self._step4_create_sections()
            steps_results["sections"] = section_result
            
            # 步骤5：创建目录
            logger.info("步骤5：创建目录")
            toc_result = self._step5_create_toc()
            steps_results["toc"] = toc_result
            
            # 步骤6：保存文档
            logger.info("步骤6：保存文档")
            output_path = self._step6_save_document(source_file_path)
            steps_results["save"] = {"success": True, "output_path": output_path}
            
            # 生成处理报告
            report = self._generate_detailed_report(steps_results)
            
            return {
                "success": True,
                "output_path": output_path,
                "steps_results": steps_results,
                "report": report,
                "message": "文档格式化成功完成"
            }
            
        except Exception as e:
            logger.error(f"格式化文档失败: {str(e)}")
            return {
                "success": False,
                "error": f"格式化失败: {str(e)}"
            }
    
    def _prepare_formatting(self, source_file_path: str, source_file_id: str, format_file_id: str) -> Dict[str, Any]:
        """准备格式化所需的数据"""
        # 获取段落分析结果
        paragraph_analysis = self.analysis_storage.get_latest_paragraph_analysis(source_file_id)
        if not paragraph_analysis:
            return {
                "success": False,
                "error": "未找到段落分析结果，请先进行段落分析"
            }
        
        # 获取格式配置
        format_config_data = self.analysis_storage.get_latest_format_config(format_file_id)
        if not format_config_data:
            return {
                "success": False,
                "error": "未找到格式配置，请先生成格式配置"
            }
        
        self.format_config = format_config_data.get("format_config", {})
        self.analysis_result = paragraph_analysis.get("result", {}).get("analysis_result", [])
        
        return {"success": True}
    
    def _convert_to_docx_length(self, value: str):
        """将各种单位转换为python-docx需要的长度对象"""
        pt_value = self.unit_converter.convert_to_pt(value)
        if pt_value is not None:
            return Pt(pt_value)
        try:
            return Pt(float(value))
        except:
            return Pt(0)
    
    def _step1_apply_page_settings(self) -> Dict[str, Any]:
        """步骤1：应用页面设置"""
        result = {"success": True, "applied": {}}
        
        page_settings = self.format_config.get("page_settings", {})
        if not page_settings:
            return {"success": True, "applied": {}, "message": "无页面设置配置"}
        
        # 应用边距
        margins = page_settings.get("margins", {})
        if margins:
            for section in self.doc.sections:
                if margins.get("top"):
                    section.top_margin = self._convert_to_docx_length(margins["top"])
                if margins.get("bottom"):
                    section.bottom_margin = self._convert_to_docx_length(margins["bottom"])
                if margins.get("left"):
                    section.left_margin = self._convert_to_docx_length(margins["left"])
                if margins.get("right"):
                    section.right_margin = self._convert_to_docx_length(margins["right"])
            result["applied"]["margins"] = margins
        
        # 应用页面方向和大小
        if page_settings.get("orientation"):
            # TODO: 实现页面方向设置
            result["applied"]["orientation"] = page_settings["orientation"]
        
        if page_settings.get("size"):
            # TODO: 实现页面大小设置
            result["applied"]["size"] = page_settings["size"]
        
        return result
    
    def _step2_create_styles(self) -> Dict[str, Any]:
        """步骤2：创建或更新样式（参考create_styles_from_config方法）"""
        result = {"success": True, "created": [], "updated": [], "failed": [], "styles_mapping": {}}
        
        styles_config = self.format_config.get("styles", {})
        
        for style_key, style_config in styles_config.items():
            try:
                style_name = style_config.get("name", style_key)
                
                # 检查样式是否已存在
                existing_style = None
                try:
                    existing_style = self.doc.styles[style_name]
                    logger.info(f"样式 {style_name} 已存在，将更新配置")
                    result["updated"].append(style_name)
                except KeyError:
                    # 样式不存在，创建新样式
                    existing_style = self.doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                    result["created"].append(style_name)
                    logger.info(f"创建新样式: {style_name}")
                
                # 应用字体设置（增强版，支持中英文混合字体）
                font_config = style_config.get("font", {})
                if font_config:
                    success = self._apply_enhanced_font_settings(existing_style, font_config)
                    if not success:
                        logger.warning(f"样式 {style_name} 字体设置部分失败")
                
                # 应用段落设置（增强版，支持悬挂缩进）
                para_config = style_config.get("paragraph", {})
                if para_config:
                    success = self._apply_enhanced_paragraph_settings(existing_style, para_config)
                    if not success:
                        logger.warning(f"样式 {style_name} 段落设置部分失败")
                
                # 设置大纲级别
                outline_level = style_config.get("outline_level")
                if outline_level is not None:
                    try:
                        self._set_outline_level(existing_style, outline_level)
                        logger.debug(f"样式 {style_name} 设置大纲级别: {outline_level}")
                    except Exception as e:
                        logger.warning(f"设置样式 {style_name} 的大纲级别失败: {str(e)}")
                
                # 记录样式映射
                result["styles_mapping"][style_key] = style_name
                
            except Exception as e:
                result["failed"].append({"style": style_key, "error": str(e)})
                logger.error(f"创建/更新样式 {style_key} 失败: {str(e)}")
        
        return result
    
    def _step3_apply_styles(self) -> Dict[str, Any]:
        """步骤3：应用样式到段落"""
        result = {"success": True, "applied": 0, "skipped": 0, "failed": 0}
        
        # 创建段落号到类型的映射
        para_type_map = {item["paragraph_number"]: item["type"] for item in self.analysis_result}
        styles_config = self.format_config.get("styles", {})
        
        # 遍历文档段落
        for i, paragraph in enumerate(self.doc.paragraphs, 1):
            para_type = para_type_map.get(i)
            if para_type and para_type in styles_config:
                try:
                    style_name = styles_config[para_type].get("name", para_type)
                    paragraph.style = style_name
                    result["applied"] += 1
                    logger.debug(f"段落 {i} 应用样式: {style_name} (类型: {para_type})")
                except Exception as e:
                    result["failed"] += 1
                    logger.warning(f"段落 {i} 应用样式失败: {str(e)}")
            else:
                result["skipped"] += 1
        
        return result
    
    def _step4_create_sections(self) -> Dict[str, Any]:
        """步骤4：创建章节（处理分节、页码、页眉页脚）"""
        result = {"success": True, "sections_created": 0, "page_numbers": False, "headers_footers": False}
        
        # 处理分节符
        section_breaks = self.format_config.get("section_breaks", {})
        if section_breaks:
            # TODO: 实现分节符处理
            pass
        
        # 处理页码
        page_numbering = self.format_config.get("page_numbering", {})
        if page_numbering:
            try:
                self._add_page_numbers(page_numbering)
                result["page_numbers"] = True
            except Exception as e:
                logger.warning(f"添加页码失败: {str(e)}")
        
        # 处理页眉页脚
        headers_footers = self.format_config.get("headers_footers", {})
        if headers_footers:
            try:
                self._add_headers_footers(headers_footers)
                result["headers_footers"] = True
            except Exception as e:
                logger.warning(f"添加页眉页脚失败: {str(e)}")
        
        return result
    
    def _step5_create_toc(self) -> Dict[str, Any]:
        """步骤5：创建目录"""
        result = {"success": True, "toc_created": False}
        
        toc_settings = self.format_config.get("toc_settings", {})
        if toc_settings and toc_settings.get("auto_generate_toc", False):
            try:
                # 在文档开头添加目录
                self._insert_toc(toc_settings)
                result["toc_created"] = True
            except Exception as e:
                logger.warning(f"创建目录失败: {str(e)}")
                result["error"] = str(e)
        
        return result
    
    def _step6_save_document(self, source_path: str) -> str:
        """步骤6：保存文档"""
        base_path = Path(source_path)
        output_path = base_path.parent / f"{base_path.stem}_formatted{base_path.suffix}"
        print(f"保存到{output_path}")
        self.doc.save(str(output_path))
        return str(output_path)
    
    def _get_or_create_style(self, style_id: str, style_name: str):
        """获取或创建样式"""
        try:
            return self.doc.styles[style_name]
        except KeyError:
            return self.doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    
    def _apply_font_settings(self, style, font_config: Dict[str, Any]):
        """应用字体设置（旧版方法，保留兼容性）"""
        font = style.font
        
        if font_config.get("size"):
            size_pt = self.unit_converter.convert_to_pt(font_config["size"])
            font.size = Pt(size_pt)
        
        if font_config.get("chinese"):
            font.name = font_config["chinese"]
        
        if "bold" in font_config:
            font.bold = font_config["bold"]
        
        if "italic" in font_config:
            font.italic = font_config["italic"]
    
    def _apply_enhanced_font_settings(self, style, font_config: Dict[str, Any]) -> bool:
        """应用增强版字体设置（支持中英文混合字体）"""
        try:
            font = style.font
            
            # 基本字体设置
            chinese_font = font_config.get("chinese", "宋体")
            english_font = font_config.get("english", "Times New Roman")
            font_size = font_config.get("size", "12pt")
            
            # 设置英文字体为基础字体
            font.name = english_font
            
            # 设置字体大小
            if font_size:
                size_pt = self.unit_converter.convert_to_pt(font_size)
                if size_pt:
                    font.size = Pt(size_pt)
            
            # 设置粗体和斜体
            if "bold" in font_config:
                font.bold = font_config["bold"]
            if "italic" in font_config:
                font.italic = font_config["italic"]
            
            # 设置中英文混合字体
            self._set_mixed_font_for_style(style, chinese_font, english_font)
            
            return True
        except Exception as e:
            logger.error(f"应用增强字体设置失败: {str(e)}")
            return False
    
    def _apply_paragraph_settings(self, style, para_config: Dict[str, Any]):
        """应用段落设置"""
        para_format = style.paragraph_format
        
        # 对齐方式
        alignment = para_config.get("alignment")
        if alignment:
            alignment_map = {
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
            }
            para_format.alignment = alignment_map.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)
        
        # 行距
        if para_config.get("line_spacing"):
            line_spacing = self.unit_converter.convert_to_pt(para_config["line_spacing"])
            para_format.line_spacing = Pt(line_spacing)
        
        # 段前段后
        if para_config.get("space_before"):
            para_format.space_before = Pt(self.unit_converter.convert_to_pt(para_config["space_before"]))
        if para_config.get("space_after"):
            para_format.space_after = Pt(self.unit_converter.convert_to_pt(para_config["space_after"]))
        
        # 缩进
        if para_config.get("first_line_indent"):
            para_format.first_line_indent = self._convert_to_docx_length(para_config["first_line_indent"])
        if para_config.get("left_indent"):
            para_format.left_indent = self._convert_to_docx_length(para_config["left_indent"])
        if para_config.get("right_indent"):
            para_format.right_indent = self._convert_to_docx_length(para_config["right_indent"])
    
    def _apply_enhanced_paragraph_settings(self, style, para_config: Dict[str, Any]) -> bool:
        """应用增强版段落设置（支持悬挂缩进等高级功能）"""
        try:
            para_format = style.paragraph_format
            
            # 对齐方式
            alignment = para_config.get("alignment", "left")
            alignment_map = {
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
            }
            para_format.alignment = alignment_map.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)
            
            # 行距
            line_spacing = para_config.get("line_spacing")
            if line_spacing:
                line_spacing_pt = self.unit_converter.convert_to_pt(line_spacing)
                if line_spacing_pt:
                    para_format.line_spacing = Pt(line_spacing_pt)
            
            # 段前距
            space_before = para_config.get("space_before")
            if space_before and space_before != "0pt":
                space_before_pt = self.unit_converter.convert_to_pt(space_before)
                if space_before_pt:
                    para_format.space_before = Pt(space_before_pt)
            
            # 段后距
            space_after = para_config.get("space_after")
            if space_after and space_after != "0pt":
                space_after_pt = self.unit_converter.convert_to_pt(space_after)
                if space_after_pt:
                    para_format.space_after = Pt(space_after_pt)
            
            # 首行缩进
            first_line_indent = para_config.get("first_line_indent")
            if first_line_indent and first_line_indent != "0pt":
                first_line_pt = self.unit_converter.convert_to_pt(first_line_indent)
                if first_line_pt:
                    para_format.first_line_indent = Pt(first_line_pt)
            
            # 左缩进
            left_indent = para_config.get("left_indent")
            if left_indent and left_indent != "0pt":
                left_indent_pt = self.unit_converter.convert_to_pt(left_indent)
                if left_indent_pt:
                    para_format.left_indent = Pt(left_indent_pt)
            
            # 右缩进
            right_indent = para_config.get("right_indent")
            if right_indent and right_indent != "0pt":
                right_indent_pt = self.unit_converter.convert_to_pt(right_indent)
                if right_indent_pt:
                    para_format.right_indent = Pt(right_indent_pt)
            
            # 悬挂缩进（新功能）
            hanging_indent = para_config.get("hanging_indent")
            if hanging_indent and hanging_indent != "0pt":
                hanging_pt = self.unit_converter.convert_to_pt(hanging_indent)
                if hanging_pt:
                    # 悬挂缩进实现为负的首行缩进
                    para_format.first_line_indent = Pt(-hanging_pt)
            
            return True
        except Exception as e:
            logger.error(f"应用增强段落设置失败: {str(e)}")
            return False
    
    def _add_page_numbers(self, page_numbering: Dict[str, Any]):
        """添加页码"""
        # TODO: 实现页码添加
        # 这需要操作XML来实现
        pass
    
    def _add_headers_footers(self, headers_footers: Dict[str, Any]):
        """添加页眉页脚"""
        for section in self.doc.sections:
            # 设置奇偶页不同
            if headers_footers.get("odd_even_different"):
                section.different_first_page_header_footer = True
            
            # TODO: 实现具体的页眉页脚内容
            pass
    
    def _insert_toc(self, toc_settings: Dict[str, Any]):
        """插入目录"""
        # 在文档开头插入一个段落作为目录
        toc_para = self.doc.paragraphs[0].insert_paragraph_before()
        toc_para.text = toc_settings.get("title", "目录")
        toc_para.style = "TOCTitle" if "TOCTitle" in self.doc.styles else "Heading1"
        
        # 添加分页符
        from docx.enum.text import WD_BREAK
        toc_para.add_run().add_break(WD_BREAK.PAGE)
        
        # TODO: 实现自动目录生成
        # 这需要使用域代码实现
    
    def _generate_detailed_report(self, steps_results: Dict[str, Any]) -> Dict[str, Any]:
        """生成详细的处理报告"""
        report = {
            "total_steps": 6,
            "completed_steps": sum(1 for step in steps_results.values() if step.get("success", False)),
            "details": steps_results,
            "summary": {
                "page_settings": steps_results.get("page_settings", {}).get("applied", {}),
                "styles_created": len(steps_results.get("create_styles", {}).get("created", [])),
                "styles_updated": len(steps_results.get("create_styles", {}).get("updated", [])),
                "paragraphs_styled": steps_results.get("apply_styles", {}).get("applied", 0),
                "toc_created": steps_results.get("toc", {}).get("toc_created", False),
                "page_numbers_added": steps_results.get("sections", {}).get("page_numbers", False),
                "headers_footers_added": steps_results.get("sections", {}).get("headers_footers", False)
            }
        }
        return report
    
    def _set_outline_level(self, style, outline_level: int):
        """设置样式的大纲级别"""
        try:
            # 获取样式的XML元素
            style_element = style._element
            
            # 查找或创建pPr元素
            pPr = style_element.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                style_element.append(pPr)
            
            # 查找或创建outlineLvl元素
            outline_elem = pPr.find(qn('w:outlineLvl'))
            if outline_elem is None:
                outline_elem = OxmlElement('w:outlineLvl')
                pPr.append(outline_elem)
            
            # 设置大纲级别值
            outline_elem.set(qn('w:val'), str(outline_level))
            
        except Exception as e:
            logger.warning(f"设置大纲级别失败: {str(e)}")
            raise
    
    def _set_mixed_font_for_style(self, style, chinese_font='宋体', english_font='Times New Roman'):
        """为样式设置中英文混合字体"""
        try:
            # 设置基础字体（英文）
            style.font.name = english_font
            
            # 设置东亚字体（中文）
            style_element = style._element
            rpr = style_element.find(qn('w:rPr'))
            if rpr is None:
                rpr = OxmlElement('w:rPr')
                style_element.append(rpr)
            
            # 查找或创建rFonts元素
            rfonts = rpr.find(qn('w:rFonts'))
            if rfonts is None:
                rfonts = OxmlElement('w:rFonts')
                rpr.append(rfonts)
            
            # 设置各种字体类型
            rfonts.set(qn('w:eastAsia'), chinese_font)    # 东亚字体（中文）
            rfonts.set(qn('w:ascii'), english_font)       # ASCII字体（英文）
            rfonts.set(qn('w:hAnsi'), english_font)       # 高ANSI字体（英文）
            
            return True
        except Exception as e:
            logger.error(f"设置混合字体失败: {str(e)}")
            return False