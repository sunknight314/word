#!/usr/bin/env python3
"""
配置加载器 - 加载和解析文档修改配置
"""

import json
import os
from typing import Dict, Any, Optional, Union
from docx.shared import Pt, Inches, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION, WD_ORIENT
from enum import Enum

class ModifyMode(Enum):
    """修改模式枚举"""
    MERGE = "merge"        # 合并模式：保留原有+应用新的
    REPLACE = "replace"    # 替换模式：完全替换
    APPEND = "append"      # 追加模式：在原有基础上追加
    SELECTIVE = "selective" # 选择性模式：只修改指定部分

class ModifyConfigLoader:
    """文档修改配置加载器"""
    
    def __init__(self, config_path: str):
        """初始化配置加载器"""
        self.config_path = config_path
        self.config: Dict[str, Any] = {}
        self.load_config()
    
    def load_config(self) -> None:
        """加载配置文件"""
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(f"配置文件不存在: {self.config_path}")
        
        with open(self.config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        print(f"✅ 成功加载配置文件: {self.config_path}")
    
    # ==================== 获取配置部分 ====================
    
    def get_modify_mode(self) -> ModifyMode:
        """获取修改模式"""
        mode_str = self.config.get('modify_mode', 'merge')
        return ModifyMode(mode_str)
    
    def get_document_info(self) -> Dict[str, Any]:
        """获取文档信息配置"""
        return self.config.get('document_info', {})
    
    def get_page_settings(self) -> Dict[str, Any]:
        """获取页面设置配置"""
        return self.config.get('page_settings', {})
    
    def get_styles_config(self) -> Dict[str, Any]:
        """获取样式配置"""
        return self.config.get('styles', {})
    
    def get_structure_config(self) -> Dict[str, Any]:
        """获取文档结构配置"""
        return self.config.get('document_structure', {})
    
    def get_content_modifications(self) -> Dict[str, Any]:
        """获取内容修改配置"""
        return self.config.get('content_modifications', {})
    
    def get_page_numbering_config(self) -> Dict[str, Any]:
        """获取页码配置"""
        return self.config.get('page_numbering', {})
    
    def get_headers_footers_config(self) -> Dict[str, Any]:
        """获取页眉页脚配置"""
        return self.config.get('headers_footers', {})
    
    def get_section_breaks_config(self) -> Dict[str, Any]:
        """获取分节符配置"""
        return self.config.get('section_breaks', {})
    
    def get_toc_settings(self) -> Dict[str, Any]:
        """获取目录设置"""
        return self.config.get('toc_settings', {})
    
    # ==================== 解析工具方法 ====================
    
    @staticmethod
    def parse_length(value: Union[str, int, float]) -> Optional[Union[Pt, Inches, Cm, Mm]]:
        """解析长度值"""
        if value is None:
            return None
        
        if isinstance(value, (int, float)):
            return Pt(value)
        
        value_str = str(value).strip().lower()
        
        # 解析不同单位
        if value_str.endswith('pt'):
            return Pt(float(value_str[:-2]))
        elif value_str.endswith('in') or value_str.endswith('inch'):
            return Inches(float(value_str.replace('inch', '').replace('in', '').strip()))
        elif value_str.endswith('cm'):
            return Cm(float(value_str[:-2]))
        elif value_str.endswith('mm'):
            return Mm(float(value_str[:-2]))
        else:
            # 默认为磅
            try:
                return Pt(float(value_str))
            except ValueError:
                return None
    
    @staticmethod
    def parse_alignment(align_str: str) -> Optional[WD_ALIGN_PARAGRAPH]:
        """解析对齐方式"""
        if not align_str:
            return None
        
        align_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
            'distribute': WD_ALIGN_PARAGRAPH.DISTRIBUTE,
            'justify_low': WD_ALIGN_PARAGRAPH.JUSTIFY_LOW,
            'justify_med': WD_ALIGN_PARAGRAPH.JUSTIFY_MED,
            'justify_hi': WD_ALIGN_PARAGRAPH.JUSTIFY_HI,
            'thai_justify': WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
        }
        
        return align_map.get(align_str.lower())
    
    @staticmethod
    def parse_line_spacing_rule(rule_str: str) -> Optional[WD_LINE_SPACING]:
        """解析行距规则"""
        if not rule_str:
            return None
        
        rule_map = {
            'single': WD_LINE_SPACING.SINGLE,
            'one_point_five': WD_LINE_SPACING.ONE_POINT_FIVE,
            '1.5': WD_LINE_SPACING.ONE_POINT_FIVE,
            'double': WD_LINE_SPACING.DOUBLE,
            'at_least': WD_LINE_SPACING.AT_LEAST,
            'exactly': WD_LINE_SPACING.EXACTLY,
            'multiple': WD_LINE_SPACING.MULTIPLE
        }
        
        return rule_map.get(rule_str.lower())
    
    @staticmethod
    def parse_orientation(orient_str: str) -> Optional[WD_ORIENT]:
        """解析页面方向"""
        if not orient_str:
            return None
        
        orient_map = {
            'portrait': WD_ORIENT.PORTRAIT,
            'landscape': WD_ORIENT.LANDSCAPE
        }
        
        return orient_map.get(orient_str.lower())
    
    @staticmethod
    def parse_section_break_type(break_type: str) -> Optional[WD_SECTION]:
        """解析分节符类型"""
        if not break_type:
            return None
        
        break_map = {
            'continuous': WD_SECTION.CONTINUOUS,
            'new_column': WD_SECTION.NEW_COLUMN,
            'new_page': WD_SECTION.NEW_PAGE,
            'even_page': WD_SECTION.EVEN_PAGE,
            'odd_page': WD_SECTION.ODD_PAGE
        }
        
        return break_map.get(break_type.lower())
    
    @staticmethod
    def parse_color(color_str: str) -> Optional[tuple]:
        """解析颜色值（返回RGB元组）"""
        if not color_str:
            return None
        
        # 移除#号
        color_str = color_str.strip('#')
        
        # 解析16进制颜色
        if len(color_str) == 6:
            try:
                r = int(color_str[0:2], 16)
                g = int(color_str[2:4], 16)
                b = int(color_str[4:6], 16)
                return (r, g, b)
            except ValueError:
                return None
        
        # 预定义颜色
        color_map = {
            'black': (0, 0, 0),
            'white': (255, 255, 255),
            'red': (255, 0, 0),
            'green': (0, 255, 0),
            'blue': (0, 0, 255),
            'gray': (128, 128, 128),
            'dark_gray': (64, 64, 64),
            'light_gray': (192, 192, 192)
        }
        
        return color_map.get(color_str.lower())
    
    @staticmethod
    def parse_boolean(value: Any) -> bool:
        """解析布尔值"""
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.lower() in ('true', 'yes', '1', 'on')
        return bool(value)
    
    # ==================== 验证方法 ====================
    
    def validate_config(self) -> bool:
        """验证配置文件的有效性"""
        required_fields = []
        errors = []
        
        # 检查必需字段
        for field in required_fields:
            if field not in self.config:
                errors.append(f"缺少必需字段: {field}")
        
        # 检查修改模式
        mode_str = self.config.get('modify_mode', 'merge')
        try:
            ModifyMode(mode_str)
        except ValueError:
            errors.append(f"无效的修改模式: {mode_str}")
        
        if errors:
            print("❌ 配置验证失败:")
            for error in errors:
                print(f"  - {error}")
            return False
        
        return True
    
    def get_style_mapping(self) -> Dict[str, str]:
        """获取样式映射配置"""
        return self.config.get('style_mapping', {})
    
    def get_batch_operations(self) -> list:
        """获取批量操作配置"""
        return self.config.get('batch_operations', [])
    
    def should_preserve_formatting(self) -> bool:
        """是否保留原有格式"""
        return self.config.get('preserve_formatting', False)
    
    def get_selective_modifications(self) -> Dict[str, Any]:
        """获取选择性修改配置"""
        return self.config.get('selective_modifications', {})

def main():
    """测试配置加载器"""
    
    # 创建测试配置文件
    test_config = {
        "modify_mode": "merge",
        "document_info": {
            "title": "修改后的文档",
            "author": "配置修改器",
            "subject": "基于配置的文档修改"
        },
        "page_settings": {
            "page_size": "A4",
            "orientation": "portrait",
            "margins": {
                "top": "2.54cm",
                "bottom": "2.54cm",
                "left": "3.18cm",
                "right": "3.18cm"
            }
        },
        "styles": {
            "Heading 1": {
                "font": {
                    "name": "微软雅黑",
                    "size": "16pt",
                    "bold": true,
                    "color": "#000080"
                },
                "paragraph": {
                    "alignment": "left",
                    "space_before": "12pt",
                    "space_after": "6pt"
                }
            }
        },
        "preserve_formatting": false
    }
    
    # 保存测试配置
    config_path = "test_modify_config.json"
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(test_config, f, ensure_ascii=False, indent=2)
    
    # 测试加载器
    loader = ModifyConfigLoader(config_path)
    
    print(f"\n修改模式: {loader.get_modify_mode()}")
    print(f"文档信息: {loader.get_document_info()}")
    print(f"页面设置: {loader.get_page_settings()}")
    
    # 测试解析方法
    print(f"\n解析测试:")
    print(f"解析长度 '12pt': {loader.parse_length('12pt')}")
    print(f"解析长度 '2.54cm': {loader.parse_length('2.54cm')}")
    print(f"解析对齐 'center': {loader.parse_alignment('center')}")
    print(f"解析颜色 '#000080': {loader.parse_color('#000080')}")
    
    # 清理测试文件
    os.remove(config_path)
    print(f"\n✅ 配置加载器测试完成")

if __name__ == "__main__":
    main()