"""
文档格式转换服务
根据段落分析结果和格式配置对文档进行格式化
"""
import os
from typing import Dict, Any, List, Optional
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from app.services.analysis_result_storage import AnalysisResultStorage
from app.utils.unit_converter import UnitConverter
from app.constants.style_names import StyleNames, get_style_description
import logging

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class DocumentFormatter:
    """文档格式化类"""
    
    def __init__(self):
        self.analysis_storage = AnalysisResultStorage()
        self.unit_converter = UnitConverter()
    
    def _convert_to_docx_length(self, value: str):
        """将各种单位转换为python-docx需要的长度对象"""
        pt_value = self.unit_converter.convert_to_pt(value)
        if pt_value is not None:
            return Pt(pt_value)
        # 如果无法转换，尝试直接创建Pt对象
        try:
            return Pt(float(value))
        except:
            return Pt(0)  # 默认值
        
    def format_document(self, source_file_path: str, source_file_id: str, format_file_id: str) -> Dict[str, Any]:
        """
        格式化文档
        
        Args:
            source_file_path: 源文档路径
            source_file_id: 源文档ID（用于获取段落分析结果）
            format_file_id: 格式文档ID（用于获取格式配置）
            
        Returns:
            处理结果
        """
        try:
            # 1. 获取段落分析结果
            paragraph_analysis = self.analysis_storage.get_latest_paragraph_analysis(source_file_id)
            if not paragraph_analysis:
                return {
                    "success": False,
                    "error": "未找到段落分析结果，请先进行段落分析"
                }
            
            # 2. 获取格式配置
            format_config_data = self.analysis_storage.get_latest_format_config(format_file_id)
            if not format_config_data:
                return {
                    "success": False,
                    "error": "未找到格式配置，请先生成格式配置"
                }
            
            format_config = format_config_data.get("format_config", {})
            
            # 3. 打开源文档
            doc = Document(source_file_path)
            
            # 4. 应用页面设置
            self._apply_page_settings(doc, format_config.get("page_settings", {}))
            
            # 5. 创建或更新样式
            styles_config = format_config.get("styles", {})
            self._update_document_styles(doc, styles_config)
            
            # 6. 应用样式到段落
            analysis_result = paragraph_analysis.get("result", {}).get("analysis_result", [])
            self._apply_styles_to_paragraphs(doc, analysis_result, styles_config)
            
            # 7. 保存格式化后的文档
            output_path = self._get_output_path(source_file_path)
            doc.save(output_path)
            
            # 8. 生成处理报告
            report = self._generate_format_report(analysis_result, styles_config)
            
            return {
                "success": True,
                "output_path": output_path,
                "report": report,
                "message": "文档格式化成功"
            }
            
        except Exception as e:
            logger.error(f"格式化文档失败: {str(e)}")
            return {
                "success": False,
                "error": f"格式化失败: {str(e)}"
            }
    
    def _apply_page_settings(self, doc: Document, page_settings: Dict[str, Any]):
        """应用页面设置"""
        if not page_settings:
            return
            
        # 应用边距
        margins = page_settings.get("margins", {})
        for section in doc.sections:
            if margins.get("top"):
                section.top_margin = self._convert_to_docx_length(margins["top"])
            if margins.get("bottom"):
                section.bottom_margin = self._convert_to_docx_length(margins["bottom"])
            if margins.get("left"):
                section.left_margin = self._convert_to_docx_length(margins["left"])
            if margins.get("right"):
                section.right_margin = self._convert_to_docx_length(margins["right"])
                
        logger.info(f"已应用页面设置: {page_settings}")
    
    def _update_document_styles(self, doc: Document, styles_config: Dict[str, Any]):
        """创建或更新文档样式"""
        for style_name, style_config in styles_config.items():
            try:
                # 获取或创建样式
                style = self._get_or_create_style(doc, style_name, style_config.get("name", style_name))
                
                # 应用字体设置
                font_config = style_config.get("font", {})
                if font_config:
                    self._apply_font_settings(style, font_config)
                
                # 应用段落设置
                para_config = style_config.get("paragraph", {})
                if para_config:
                    self._apply_paragraph_settings(style, para_config)
                    
                logger.info(f"已更新样式: {style_name}")
                
            except Exception as e:
                logger.warning(f"更新样式 {style_name} 失败: {str(e)}")
    
    def _get_or_create_style(self, doc: Document, style_id: str, style_name: str):
        """获取或创建样式"""
        # 尝试获取现有样式
        try:
            return doc.styles[style_name]
        except KeyError:
            # 创建新样式
            return doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    
    def _apply_font_settings(self, style, font_config: Dict[str, Any]):
        """应用字体设置"""
        font = style.font
        
        # 字号
        if font_config.get("size"):
            size_pt = self.unit_converter.convert_to_pt(font_config["size"])
            font.size = Pt(size_pt)
        
        # 字体名称
        if font_config.get("chinese"):
            font.name = font_config["chinese"]
        
        # 加粗
        if "bold" in font_config:
            font.bold = font_config["bold"]
        
        # 斜体
        if "italic" in font_config:
            font.italic = font_config["italic"]
    
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
    
    def _apply_styles_to_paragraphs(self, doc: Document, analysis_result: List[Dict], styles_config: Dict):
        """将样式应用到对应的段落"""
        # 创建段落号到类型的映射
        para_type_map = {item["paragraph_number"]: item["type"] for item in analysis_result}
        
        # 遍历文档段落
        for i, paragraph in enumerate(doc.paragraphs, 1):
            para_type = para_type_map.get(i)
            if para_type and para_type in styles_config:
                try:
                    style_name = styles_config[para_type].get("name", para_type)
                    paragraph.style = style_name
                    logger.info(f"段落 {i} 应用样式: {style_name} (类型: {para_type})")
                except Exception as e:
                    logger.warning(f"段落 {i} 应用样式失败: {str(e)}")
    
    def _get_output_path(self, source_path: str) -> str:
        """生成输出文件路径"""
        base_path = Path(source_path)
        output_path = base_path.parent / f"{base_path.stem}_formatted{base_path.suffix}"
        return str(output_path)
    
    def _generate_format_report(self, analysis_result: List[Dict], styles_config: Dict) -> Dict[str, Any]:
        """生成格式化报告"""
        # 统计各类型段落数量
        type_counts = {}
        for item in analysis_result:
            para_type = item["type"]
            type_counts[para_type] = type_counts.get(para_type, 0) + 1
        
        # 统计已定义样式的段落
        styled_count = sum(count for para_type, count in type_counts.items() if para_type in styles_config)
        total_count = sum(type_counts.values())
        
        return {
            "total_paragraphs": total_count,
            "styled_paragraphs": styled_count,
            "type_distribution": type_counts,
            "undefined_styles": [t for t in type_counts.keys() if t not in styles_config]
        }