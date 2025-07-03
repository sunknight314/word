"""
格式文件JSON结构定义
用于标准化Word文档样式信息的表示
"""

from typing import Dict, List, Any, Optional
from pydantic import BaseModel

class Margins(BaseModel):
    """页边距设置"""
    top: float = 2.54  # 厘米
    bottom: float = 2.54
    left: float = 3.18
    right: float = 3.18

class PageSetup(BaseModel):
    """页面设置"""
    margins: Margins
    orientation: str = "portrait"  # portrait/landscape
    paper_size: str = "A4"
    page_width: Optional[float] = None  # 厘米
    page_height: Optional[float] = None

class FontStyle(BaseModel):
    """字体样式"""
    name: str = "宋体"
    size: float = 12  # 磅
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Optional[str] = None  # RGB hex值

class ParagraphStyle(BaseModel):
    """段落样式"""
    name: str
    font: FontStyle
    alignment: str = "left"  # left/center/right/justify
    line_spacing: float = 1.0  # 行间距倍数
    space_before: float = 0  # 段前间距（磅）
    space_after: float = 0   # 段后间距（磅）
    left_indent: float = 0   # 左缩进（厘米）
    right_indent: float = 0  # 右缩进（厘米）
    first_line_indent: float = 0  # 首行缩进（厘米）
    keep_with_next: bool = False  # 与下段同页
    keep_together: bool = False   # 段中不分页

class CharacterStyle(BaseModel):
    """字符样式"""
    name: str
    font: FontStyle
    highlight_color: Optional[str] = None

class TableStyle(BaseModel):
    """表格样式"""
    name: str
    border_style: str = "single"  # none/single/double/dotted等
    border_width: float = 0.5  # 磅
    border_color: str = "#000000"
    cell_margins: Dict[str, float] = {"top": 0, "bottom": 0, "left": 0.19, "right": 0.19}
    
class ListStyle(BaseModel):
    """列表样式"""
    name: str
    list_type: str = "bullet"  # bullet/number
    levels: List[Dict[str, Any]] = []  # 各级别设置

class DocumentFormat(BaseModel):
    """完整文档格式定义"""
    document_name: str
    page_setup: PageSetup
    paragraph_styles: Dict[str, ParagraphStyle]
    character_styles: Dict[str, CharacterStyle] = {}
    table_styles: Dict[str, TableStyle] = {}
    list_styles: Dict[str, ListStyle] = {}
    custom_rules: List[str] = []  # 自定义格式规则描述

# 预定义的标准样式模板
STANDARD_STYLES = {
    "title": ParagraphStyle(
        name="title",
        font=FontStyle(name="宋体", size=16, bold=True),
        alignment="center",
        space_after=12
    ),
    "heading1": ParagraphStyle(
        name="heading1", 
        font=FontStyle(name="宋体", size=14, bold=True),
        alignment="left",
        space_before=12,
        space_after=6
    ),
    "heading2": ParagraphStyle(
        name="heading2",
        font=FontStyle(name="宋体", size=13, bold=True),
        alignment="left", 
        space_before=6,
        space_after=3
    ),
    "normal": ParagraphStyle(
        name="normal",
        font=FontStyle(name="宋体", size=12),
        alignment="justify",
        line_spacing=1.5,
        first_line_indent=2.0  # 首行缩进2字符
    )
}