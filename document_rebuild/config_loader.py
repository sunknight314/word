"""
配置加载器 - 从JSON文件加载格式配置
"""

import json
import os
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION

class FormatConfigLoader:
    """格式配置加载器"""
    
    def __init__(self, config_path="format_config.json"):
        """初始化配置加载器"""
        self.config_path = config_path
        self.config = None
        
    def load_config(self):
        """加载配置文件"""
        try:
            if not os.path.exists(self.config_path):
                print(f"❌ 配置文件不存在: {self.config_path}")
                return None
                
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
                
            print(f"✅ 成功加载格式配置: {self.config_path}")
            return self.config
            
        except Exception as e:
            print(f"❌ 加载配置文件失败: {e}")
            return None
    
    def get_document_info(self):
        """获取文档信息"""
        if not self.config:
            return {}
        return self.config.get("document_info", {})
    
    def get_page_settings(self):
        """获取页面设置"""
        if not self.config:
            return {}
        return self.config.get("page_settings", {})
    
    def get_styles_config(self):
        """获取样式配置"""
        if not self.config:
            return {}
        return self.config.get("styles", {})
    
    def get_toc_config(self):
        """获取目录配置"""
        if not self.config:
            return {}
        return self.config.get("toc_settings", {})
    
    def get_page_numbering_config(self):
        """获取页码配置"""
        if not self.config:
            return {}
        return self.config.get("page_numbering", {})
    
    def get_headers_footers_config(self):
        """获取页眉页脚配置"""
        if not self.config:
            return {}
        return self.config.get("headers_footers", {})
    
    def get_section_breaks_config(self):
        """获取分节符配置"""
        if not self.config:
            return {}
        return self.config.get("section_breaks", {})
    
    def get_document_structure_config(self):
        """获取文档结构配置"""
        if not self.config:
            return {}
        return self.config.get("document_structure", {})

    def parse_length(self, length_str):
        """解析长度字符串为docx对象"""
        if not length_str:
            return None
            
        try:
            if length_str.endswith('pt'):
                return Pt(float(length_str[:-2]))
            elif length_str.endswith('cm'):
                return Cm(float(length_str[:-2]))
            elif length_str.endswith('in'):
                return Inches(float(length_str[:-2]))
            else:
                # 默认为磅
                return Pt(float(length_str))
        except:
            return None
    
    def parse_alignment(self, alignment_str):
        """解析对齐方式字符串"""
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        return alignment_map.get(alignment_str, WD_ALIGN_PARAGRAPH.LEFT)
    
    def parse_section_start(self, start_str):
        """解析分节符类型"""
        start_map = {
            "continuous": WD_SECTION_START.CONTINUOUS,
            "new_column": WD_SECTION_START.NEW_COLUMN,
            "new_page": WD_SECTION_START.NEW_PAGE,
            "even_page": WD_SECTION_START.EVEN_PAGE,
            "odd_page": WD_SECTION_START.ODD_PAGE
        }
        return start_map.get(start_str, WD_SECTION_START.NEW_PAGE)
    
    def parse_orientation(self, orientation_str):
        """解析页面方向"""
        if orientation_str == "landscape":
            return WD_ORIENTATION.LANDSCAPE
        else:
            return WD_ORIENTATION.PORTRAIT

def load_format_config(config_path="format_config.json"):
    """便捷函数：加载格式配置"""
    loader = FormatConfigLoader(config_path)
    return loader.load_config(), loader

def get_style_mapping_from_config(config_loader):
    """从配置获取样式映射"""
    styles_config = config_loader.get_styles_config()
    
    style_mapping = {}
    for content_type, style_config in styles_config.items():
        style_mapping[content_type] = style_config.get("name", f"Custom{content_type.title()}")
    
    return style_mapping