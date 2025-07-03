"""
格式分析器核心服务
整合样式提取和分析功能，提供统一的格式分析接口
"""

import json
import os
from typing import Dict, List, Any, Optional
from .style_extractor import StyleExtractor
from .file_storage import FileStorage

class FormatAnalyzer:
    """格式分析器"""
    
    def __init__(self):
        self.style_extractor = StyleExtractor()
        self.file_storage = FileStorage()
    
    async def analyze_format_document(self, file_path: str, use_ai: bool = False) -> Dict[str, Any]:
        """
        分析格式文档，提取样式信息
        
        Args:
            file_path: 格式文档路径
            use_ai: 是否使用AI辅助分析（暂未实现）
            
        Returns:
            格式分析结果
        """
        try:
            # 第一阶段：使用python-docx直接提取样式
            format_info = self.style_extractor.extract_styles_from_document(file_path)
            
            # 标准化格式信息
            standardized_format = self._standardize_format_info(format_info)
            
            # 生成格式分析摘要
            analysis_summary = self._generate_format_summary(standardized_format)
            
            # 构建完整结果
            result = {
                "file_name": os.path.basename(file_path),
                "analysis_method": "python-docx",
                "extraction_time": self._get_current_time(),
                "format_data": standardized_format,
                "analysis_summary": analysis_summary,
                "recommendations": self._generate_recommendations(standardized_format)
            }
            
            return result
            
        except Exception as e:
            raise Exception(f"格式分析失败: {str(e)}")
    
    def _standardize_format_info(self, raw_format: Dict[str, Any]) -> Dict[str, Any]:
        """标准化格式信息，统一数据结构"""
        
        # 映射常见样式名称到标准名称
        style_name_mapping = {
            "标题": "title",
            "Title": "title", 
            "标题 1": "heading1",
            "Heading 1": "heading1",
            "标题 2": "heading2", 
            "Heading 2": "heading2",
            "标题 3": "heading3",
            "Heading 3": "heading3",
            "标题 4": "heading4",
            "Heading 4": "heading4",
            "正文": "normal",
            "Normal": "normal",
            "Body Text": "normal"
        }
        
        standardized = {
            "document_info": {
                "name": raw_format.get("document_name", ""),
                "extraction_method": raw_format.get("extraction_method", "python-docx")
            },
            "page_setup": raw_format.get("page_setup", {}),
            "paragraph_styles": {},
            "character_styles": raw_format.get("character_styles", {}),
            "table_styles": raw_format.get("table_styles", {}),
            "content_analysis": raw_format.get("content_analysis", {})
        }
        
        # 标准化段落样式名称
        original_styles = raw_format.get("paragraph_styles", {})
        for style_name, style_data in original_styles.items():
            # 映射到标准名称
            standard_name = style_name_mapping.get(style_name, style_name.lower().replace(" ", "_"))
            standardized["paragraph_styles"][standard_name] = style_data
        
        return standardized
    
    def _generate_format_summary(self, format_data: Dict[str, Any]) -> Dict[str, Any]:
        """生成格式分析摘要"""
        paragraph_styles = format_data.get("paragraph_styles", {})
        page_setup = format_data.get("page_setup", {})
        content_analysis = format_data.get("content_analysis", {})
        
        # 统计样式信息
        style_count = len(paragraph_styles)
        
        # 分析字体使用情况
        font_families = set()
        font_sizes = []
        for style in paragraph_styles.values():
            if isinstance(style, dict) and "font" in style:
                font_families.add(style["font"]["name"])
                font_sizes.append(style["font"]["size"])
        
        # 分析页面设置
        margins = page_setup.get("margins", {})
        paper_info = {
            "size": page_setup.get("paper_size", "Unknown"),
            "orientation": page_setup.get("orientation", "portrait"),
            "margins_uniform": self._check_uniform_margins(margins)
        }
        
        # 生成格式特征描述
        format_features = self._identify_format_features(format_data)
        
        return {
            "total_styles": style_count,
            "main_fonts": list(font_families),
            "font_size_range": {
                "min": min(font_sizes) if font_sizes else 12,
                "max": max(font_sizes) if font_sizes else 12,
                "count": len(set(font_sizes))
            },
            "paper_info": paper_info,
            "format_features": format_features,
            "style_distribution": content_analysis.get("style_usage", {}),
            "format_complexity": self._assess_format_complexity(format_data)
        }
    
    def _identify_format_features(self, format_data: Dict[str, Any]) -> List[str]:
        """识别格式特征"""
        features = []
        paragraph_styles = format_data.get("paragraph_styles", {})
        
        # 检查是否有标题层级
        if any("heading" in name or "title" in name for name in paragraph_styles.keys()):
            features.append("多级标题结构")
        
        # 检查首行缩进
        for style in paragraph_styles.values():
            if isinstance(style, dict) and style.get("first_line_indent", 0) > 0:
                features.append("首行缩进")
                break
        
        # 检查行间距
        for style in paragraph_styles.values():
            if isinstance(style, dict) and style.get("line_spacing", 1.0) != 1.0:
                features.append("自定义行间距")
                break
        
        # 检查对齐方式
        alignments = set()
        for style in paragraph_styles.values():
            if isinstance(style, dict):
                alignments.add(style.get("alignment", "left"))
        
        if "center" in alignments:
            features.append("居中对齐")
        if "justify" in alignments:
            features.append("两端对齐")
        
        # 检查字体样式
        has_bold = any(
            isinstance(style, dict) and style.get("font", {}).get("bold", False)
            for style in paragraph_styles.values()
        )
        if has_bold:
            features.append("粗体样式")
        
        return features if features else ["基础格式"]
    
    def _check_uniform_margins(self, margins: Dict[str, float]) -> bool:
        """检查页边距是否统一"""
        if not margins:
            return True
        
        values = list(margins.values())
        return len(set(values)) <= 1
    
    def _assess_format_complexity(self, format_data: Dict[str, Any]) -> str:
        """评估格式复杂度"""
        style_count = len(format_data.get("paragraph_styles", {}))
        features_count = len(self._identify_format_features(format_data))
        
        if style_count <= 3 and features_count <= 2:
            return "简单"
        elif style_count <= 8 and features_count <= 5:
            return "中等"
        else:
            return "复杂"
    
    def _generate_recommendations(self, format_data: Dict[str, Any]) -> List[str]:
        """生成格式优化建议"""
        recommendations = []
        paragraph_styles = format_data.get("paragraph_styles", {})
        page_setup = format_data.get("page_setup", {})
        
        # 检查页边距
        margins = page_setup.get("margins", {})
        if margins:
            if margins.get("left", 0) < 2.0 or margins.get("right", 0) < 2.0:
                recommendations.append("建议左右页边距至少2厘米，提高可读性")
        
        # 检查字体大小
        small_fonts = []
        for name, style in paragraph_styles.items():
            if isinstance(style, dict) and style.get("font", {}).get("size", 12) < 10:
                small_fonts.append(name)
        
        if small_fonts:
            recommendations.append(f"建议增大字体大小：{', '.join(small_fonts)}")
        
        # 检查行间距
        tight_spacing = []
        for name, style in paragraph_styles.items():
            if isinstance(style, dict) and style.get("line_spacing", 1.0) < 1.2:
                tight_spacing.append(name)
        
        if tight_spacing:
            recommendations.append("建议增加行间距至1.2倍以上，提高阅读体验")
        
        # 默认建议
        if not recommendations:
            recommendations.append("格式设置合理，建议保持当前样式")
        
        return recommendations
    
    def _get_current_time(self) -> str:
        """获取当前时间字符串"""
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def export_format_template(self, format_data: Dict[str, Any], output_path: str) -> str:
        """导出格式模板为JSON文件"""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(format_data, f, ensure_ascii=False, indent=2)
            return output_path
        except Exception as e:
            raise Exception(f"导出格式模板失败: {str(e)}")
    
    def load_format_template(self, template_path: str) -> Dict[str, Any]:
        """加载格式模板"""
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            raise Exception(f"加载格式模板失败: {str(e)}")