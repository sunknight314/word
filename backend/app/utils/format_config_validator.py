"""
格式配置验证器 - 验证格式配置的合法性
"""

from typing import Dict, List, Tuple, Optional


class FormatConfigValidator:
    """格式配置验证器类"""
    
    # 验证规则
    VALIDATION_RULES = {
        "margins": {
            "min": 0,
            "max": 300,  # pt
            "warning_min": 10,  # pt
            "warning_max": 150,  # pt
            "warning_message": "页边距通常在0.35-5.3厘米之间"
        },
        "font_size": {
            "min": 5,
            "max": 100,  # pt
            "warning_min": 8,
            "warning_max": 72,
            "warning_message": "字号通常在8-72pt之间"
        },
        "line_spacing": {
            "min": 0.5,
            "max": 10,  # 倍数或pt值
            "warning_min": 1,
            "warning_max": 3,
            "warning_message": "行距通常在1-3倍之间"
        },
        "paragraph_spacing": {
            "min": 0,
            "max": 200,  # pt
            "warning_min": 0,
            "warning_max": 50,
            "warning_message": "段间距通常在0-50pt之间"
        },
        "indent": {
            "min": 0,
            "max": 200,  # pt
            "warning_min": 0,
            "warning_max": 72,
            "warning_message": "缩进通常在0-72pt之间"
        }
    }
    
    # 必需的页面设置字段
    REQUIRED_PAGE_SETTINGS = ["paper_size", "orientation", "margins"]
    
    # 必需的样式字段
    REQUIRED_STYLE_FIELDS = ["font_size", "font_name_cn", "alignment", "line_spacing"]
    
    # 支持的纸张大小
    SUPPORTED_PAPER_SIZES = ["A3", "A4", "A5", "B4", "B5", "Letter", "Legal"]
    
    # 支持的页面方向
    SUPPORTED_ORIENTATIONS = ["portrait", "landscape", "纵向", "横向"]
    
    # 支持的对齐方式
    SUPPORTED_ALIGNMENTS = ["left", "center", "right", "justify", "左对齐", "居中", "右对齐", "两端对齐"]
    
    @classmethod
    def validate_format_config(cls, config: Dict) -> Tuple[bool, List[str], List[str]]:
        """
        验证格式配置
        
        Args:
            config: 格式配置字典
            
        Returns:
            (是否有效, 错误列表, 警告列表)
        """
        errors = []
        warnings = []
        
        # 验证基本结构
        if not isinstance(config, dict):
            errors.append("格式配置必须是字典类型")
            return False, errors, warnings
        
        # 验证页面设置
        if "page_settings" in config:
            page_errors, page_warnings = cls._validate_page_settings(config["page_settings"])
            errors.extend(page_errors)
            warnings.extend(page_warnings)
        else:
            errors.append("缺少页面设置(page_settings)")
        
        # 验证样式设置
        if "styles" in config:
            style_errors, style_warnings = cls._validate_styles(config["styles"])
            errors.extend(style_errors)
            warnings.extend(style_warnings)
        else:
            errors.append("缺少样式设置(styles)")
        
        return len(errors) == 0, errors, warnings
    
    @classmethod
    def _validate_page_settings(cls, page_settings: Dict) -> Tuple[List[str], List[str]]:
        """验证页面设置"""
        errors = []
        warnings = []
        
        # 检查必需字段
        for field in cls.REQUIRED_PAGE_SETTINGS:
            if field not in page_settings:
                errors.append(f"页面设置缺少必需字段: {field}")
        
        # 验证纸张大小
        if "paper_size" in page_settings:
            if page_settings["paper_size"] not in cls.SUPPORTED_PAPER_SIZES:
                warnings.append(f"不常用的纸张大小: {page_settings['paper_size']}")
        
        # 验证页面方向
        if "orientation" in page_settings:
            if page_settings["orientation"] not in cls.SUPPORTED_ORIENTATIONS:
                errors.append(f"不支持的页面方向: {page_settings['orientation']}")
        
        # 验证页边距
        if "margins" in page_settings:
            margin_errors, margin_warnings = cls._validate_margins(page_settings["margins"])
            errors.extend(margin_errors)
            warnings.extend(margin_warnings)
        
        return errors, warnings
    
    @classmethod
    def _validate_margins(cls, margins: Dict) -> Tuple[List[str], List[str]]:
        """验证页边距"""
        errors = []
        warnings = []
        
        required_margins = ["top", "bottom", "left", "right"]
        for margin in required_margins:
            if margin not in margins:
                errors.append(f"缺少{margin}边距设置")
                continue
            
            # 验证边距值
            value = cls._extract_pt_value(margins[margin])
            if value is not None:
                rule = cls.VALIDATION_RULES["margins"]
                if value < rule["min"] or value > rule["max"]:
                    errors.append(f"{margin}边距超出范围({rule['min']}-{rule['max']}pt)")
                elif value < rule["warning_min"] or value > rule["warning_max"]:
                    warnings.append(f"{margin}边距值不常见: {value}pt。{rule['warning_message']}")
        
        return errors, warnings
    
    @classmethod
    def _validate_styles(cls, styles: Dict) -> Tuple[List[str], List[str]]:
        """验证样式设置"""
        errors = []
        warnings = []
        
        if not isinstance(styles, dict):
            errors.append("样式设置必须是字典类型")
            return errors, warnings
        
        # 验证每个样式
        for style_name, style_config in styles.items():
            if not isinstance(style_config, dict):
                errors.append(f"样式'{style_name}'的配置必须是字典类型")
                continue
            
            # 检查必需字段
            for field in cls.REQUIRED_STYLE_FIELDS:
                if field not in style_config:
                    errors.append(f"样式'{style_name}'缺少必需字段: {field}")
            
            # 验证字号
            if "font_size" in style_config:
                size_value = cls._extract_pt_value(style_config["font_size"])
                if size_value is not None:
                    rule = cls.VALIDATION_RULES["font_size"]
                    if size_value < rule["min"] or size_value > rule["max"]:
                        errors.append(f"样式'{style_name}'的字号超出范围({rule['min']}-{rule['max']}pt)")
                    elif size_value < rule["warning_min"] or size_value > rule["warning_max"]:
                        warnings.append(f"样式'{style_name}'的字号不常见: {size_value}pt。{rule['warning_message']}")
            
            # 验证对齐方式
            if "alignment" in style_config:
                if style_config["alignment"] not in cls.SUPPORTED_ALIGNMENTS:
                    errors.append(f"样式'{style_name}'的对齐方式不支持: {style_config['alignment']}")
            
            # 验证行距
            if "line_spacing" in style_config:
                spacing_value = cls._extract_spacing_value(style_config["line_spacing"])
                if spacing_value is not None:
                    rule = cls.VALIDATION_RULES["line_spacing"]
                    if spacing_value < rule["min"] or spacing_value > rule["max"]:
                        errors.append(f"样式'{style_name}'的行距超出范围({rule['min']}-{rule['max']})")
            
            # 验证段间距
            for spacing_field in ["space_before", "space_after"]:
                if spacing_field in style_config:
                    spacing_value = cls._extract_pt_value(style_config[spacing_field])
                    if spacing_value is not None:
                        rule = cls.VALIDATION_RULES["paragraph_spacing"]
                        if spacing_value < rule["min"] or spacing_value > rule["max"]:
                            errors.append(f"样式'{style_name}'的{spacing_field}超出范围({rule['min']}-{rule['max']}pt)")
            
            # 验证缩进
            if "first_line_indent" in style_config:
                indent_value = cls._extract_pt_value(style_config["first_line_indent"])
                if indent_value is not None:
                    rule = cls.VALIDATION_RULES["indent"]
                    if indent_value < rule["min"] or indent_value > rule["max"]:
                        errors.append(f"样式'{style_name}'的首行缩进超出范围({rule['min']}-{rule['max']}pt)")
        
        return errors, warnings
    
    @classmethod
    def _extract_pt_value(cls, value: str) -> Optional[float]:
        """从字符串中提取pt值"""
        if not isinstance(value, str):
            return None
        
        import re
        # 匹配数字和pt单位
        match = re.match(r'^([\d.]+)\s*pt$', value.strip())
        if match:
            try:
                return float(match.group(1))
            except ValueError:
                return None
        return None
    
    @classmethod
    def _extract_spacing_value(cls, value: str) -> Optional[float]:
        """从行距字符串中提取数值（可能是倍数或pt值）"""
        if not isinstance(value, str):
            return None
        
        import re
        value = value.strip()
        
        # 匹配纯数字（倍数）
        if re.match(r'^[\d.]+$', value):
            try:
                return float(value)
            except ValueError:
                return None
        
        # 匹配pt值
        return cls._extract_pt_value(value)
    
    @classmethod
    def validate_field_value(cls, field_type: str, value: float, unit: str = "pt") -> Tuple[bool, Optional[str]]:
        """
        验证单个字段值
        
        Args:
            field_type: 字段类型
            value: 数值
            unit: 单位
            
        Returns:
            (是否有效, 错误/警告信息)
        """
        # 将值转换为pt进行验证
        if unit != "pt":
            from app.utils.unit_converter import UnitConverter
            pt_value = UnitConverter.convert_to_pt(f"{value}{unit}")
            if pt_value is None:
                return False, "无法转换单位"
        else:
            pt_value = value
        
        # 获取验证规则
        rule_key = {
            "margin": "margins",
            "fontSize": "font_size",
            "lineSpacing": "line_spacing",
            "spacing": "paragraph_spacing",
            "indent": "indent"
        }.get(field_type)
        
        if not rule_key or rule_key not in cls.VALIDATION_RULES:
            return True, None
        
        rule = cls.VALIDATION_RULES[rule_key]
        
        # 检查范围
        if pt_value < rule["min"] or pt_value > rule["max"]:
            return False, f"值必须在{rule['min']}-{rule['max']}{unit}之间"
        
        # 检查警告范围
        if pt_value < rule["warning_min"] or pt_value > rule["warning_max"]:
            return True, rule["warning_message"]
        
        return True, None