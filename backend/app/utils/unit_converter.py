"""
单位转换工具 - 将不同单位统一转换为pt（磅）
"""

import re
from typing import Union, Optional


class UnitConverter:
    """单位转换器类"""
    
    # 中文字号到磅值的映射
    CHINESE_FONT_SIZE_MAP = {
        "初号": 42,
        "小初": 36,
        "一号": 26,
        "小一": 24,
        "二号": 22,
        "小二": 18,
        "三号": 16,
        "小三": 15,
        "四号": 14,
        "小四": 12,
        "五号": 10.5,
        "小五": 9,
        "六号": 7.5,
        "小六": 6.5,
        "七号": 5.5,
        "八号": 5
    }
    
    # 单位转换比例（转换为pt）
    UNIT_CONVERSIONS = {
        "pt": 1,          # 磅
        "磅": 1,
        "px": 0.75,       # 像素（96dpi）
        "mm": 2.834646,   # 毫米
        "cm": 28.34646,   # 厘米
        "厘米": 28.34646,
        "in": 72,         # 英寸
        "inch": 72,
        "英寸": 72,
        "pc": 12,         # pica
        "em": 12,         # 相对单位，假设基准为12pt
        "字符": 12        # 假设一个字符为当前字号大小
    }
    
    @classmethod
    def convert_to_pt(cls, value: str, base_font_size: float = 12) -> Optional[float]:
        """
        将各种单位转换为pt
        
        Args:
            value: 带单位的值，如"2.54cm"、"三号"、"1.5倍"等
            base_font_size: 基准字号（用于计算相对单位）
            
        Returns:
            转换后的pt值，无法转换返回None
        """
        if not value or not isinstance(value, str):
            return None
            
        value = value.strip()
        
        # 检查是否为中文字号
        if value in cls.CHINESE_FONT_SIZE_MAP:
            return cls.CHINESE_FONT_SIZE_MAP[value]
        
        # 处理倍行距（如"1.5倍"）
        multiple_match = re.match(r'^([\d.]+)\s*倍$', value)
        if multiple_match:
            multiple = float(multiple_match.group(1))
            return base_font_size * multiple
        
        # 处理带单位的数值
        pattern = r'^([-+]?[\d.]+)\s*([a-zA-Z\u4e00-\u9fa5]+)?$'
        match = re.match(pattern, value)
        
        if not match:
            return None
            
        num_str, unit = match.groups()
        
        try:
            num = float(num_str)
        except ValueError:
            return None
        
        # 如果没有单位，默认为pt
        if not unit:
            return num
        
        unit = unit.lower()
        
        # 特殊处理字符单位
        if unit == "字符":
            return num * base_font_size
        
        # 查找转换比例
        if unit in cls.UNIT_CONVERSIONS:
            return num * cls.UNIT_CONVERSIONS[unit]
        
        return None
    
    @classmethod
    def convert_margin(cls, margin: str) -> str:
        """
        转换页边距为pt格式字符串
        
        Args:
            margin: 原始页边距值
            
        Returns:
            转换后的字符串，如"72pt"
        """
        pt_value = cls.convert_to_pt(margin)
        if pt_value is not None:
            return f"{pt_value:.1f}pt"
        return margin  # 无法转换则返回原值
    
    @classmethod
    def convert_font_size(cls, size: str) -> str:
        """
        转换字号为pt格式字符串
        
        Args:
            size: 原始字号值
            
        Returns:
            转换后的字符串，如"12pt"
        """
        pt_value = cls.convert_to_pt(size)
        if pt_value is not None:
            return f"{pt_value:.1f}pt"
        return size  # 无法转换则返回原值
    
    @classmethod
    def convert_spacing(cls, spacing: str, base_font_size: float = 12) -> str:
        """
        转换行距/段距为pt格式字符串
        
        Args:
            spacing: 原始间距值
            base_font_size: 基准字号（用于计算倍行距）
            
        Returns:
            转换后的字符串，如"18pt"
        """
        # 处理"固定值"前缀
        if spacing.startswith("固定值"):
            spacing = spacing.replace("固定值", "").strip()
        
        pt_value = cls.convert_to_pt(spacing, base_font_size)
        if pt_value is not None:
            return f"{pt_value:.1f}pt"
        return spacing  # 无法转换则返回原值
    
    @classmethod
    def convert_indent(cls, indent: str, base_font_size: float = 12) -> str:
        """
        转换缩进为pt格式字符串
        
        Args:
            indent: 原始缩进值
            base_font_size: 基准字号（用于计算字符单位）
            
        Returns:
            转换后的字符串，如"24pt"
        """
        pt_value = cls.convert_to_pt(indent, base_font_size)
        if pt_value is not None:
            return f"{pt_value:.1f}pt"
        return indent  # 无法转换则返回原值
    
    @classmethod
    def get_chinese_font_size_name(cls, pt_size: float) -> str:
        """获取磅值对应的中文字号名称"""
        # 查找最接近的中文字号
        for chinese_name, standard_pt in cls.CHINESE_FONT_SIZE_MAP.items():
            if abs(pt_size - standard_pt) < 0.5:  # 允许0.5pt的误差
                return chinese_name
        return ""
    
    @classmethod
    def convert_from_pt(cls, pt_value: float, target_unit: str) -> float:
        """将pt值转换为其他单位"""
        target_unit = target_unit.lower().strip()
        
        if target_unit == "pt" or target_unit == "磅":
            return pt_value
        elif target_unit == "px":
            return pt_value / 0.75  # 1pt = 4/3px (at 96 DPI)
        elif target_unit == "cm" or target_unit == "厘米":
            return pt_value / 28.34646  # 1cm = 28.34646pt
        elif target_unit == "mm" or target_unit == "毫米":
            return pt_value / 2.834646  # 1mm = 2.834646pt
        elif target_unit == "inch" or target_unit == "英寸":
            return pt_value / 72  # 1inch = 72pt
        elif target_unit == "em":
            # 假设基准字号为12pt
            return pt_value / 12
        elif target_unit == "字符":
            # 假设1字符 = 12pt（小四号）
            return pt_value / 12
        else:
            return pt_value
    
    @classmethod
    def get_display_value(cls, pt_value: float, field_type: str = "general") -> dict:
        """获取用于显示的值和单位
        
        Args:
            pt_value: 磅值
            field_type: 字段类型 ('fontSize', 'margin', 'lineSpacing' 等)
            
        Returns:
            包含显示值、单位、等价描述的字典
        """
        result = {
            "pt_value": pt_value,
            "display_value": pt_value,
            "unit": "pt",
            "equivalent": ""
        }
        
        if field_type == "fontSize":
            # 字号优先显示为磅值，附带中文字号
            chinese_name = cls.get_chinese_font_size_name(pt_value)
            result["display_value"] = pt_value
            result["unit"] = "pt"
            if chinese_name:
                result["equivalent"] = f"（{chinese_name}）"
        
        elif field_type in ["margin", "spacing", "indent"]:
            # 边距、间距等优先显示为厘米
            cm_value = cls.convert_from_pt(pt_value, "cm")
            result["display_value"] = round(cm_value, 2)
            result["unit"] = "cm"
            result["equivalent"] = f"（{pt_value}pt）"
        
        elif field_type == "lineSpacing":
            # 行距特殊处理
            if pt_value % 12 == 0 and pt_value > 0:
                # 如果是12的倍数，显示为倍数
                multiple = pt_value / 12
                result["display_value"] = multiple
                result["unit"] = "倍"
                result["equivalent"] = f"（{pt_value}pt）"
            else:
                # 否则显示为磅值
                result["display_value"] = pt_value
                result["unit"] = "pt"
        
        return result