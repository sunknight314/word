"""
深度分析"我的样式"的字体设置
包括继承关系和实际使用的字体
"""

from style_extractor import WordStyleExtractor
from pathlib import Path
import json
from docx import Document
from docx.oxml.ns import qn


def analyze_my_style():
    """深度分析"我的样式"的字体信息"""
    print("🔍 深度分析\"我的样式\"")
    print("=" * 50)
    
    doc_path = "../test.docx"
    
    if not Path(doc_path).exists():
        print(f"⚠️  文档不存在: {doc_path}")
        return
    
    # 使用python-docx直接分析
    doc = Document(doc_path)
    
    # 查找"我的样式"
    my_style = None
    for style in doc.styles:
        if style.name == "我的样式":
            my_style = style
            break
    
    if not my_style:
        print("❌ 未找到\"我的样式\"")
        return
    
    print(f"✅ 找到样式: {my_style.name}")
    print(f"📋 样式类型: {my_style.type}")
    print(f"🏗️  内置样式: {my_style.builtin}")
    
    # 分析字体设置
    print(f"\n🔤 字体分析:")
    analyze_font_details(my_style)
    
    # 分析继承关系
    print(f"\n🔗 继承关系分析:")
    analyze_inheritance(my_style)
    
    # 分析XML中的字体设置
    print(f"\n🔬 XML深度分析:")
    analyze_xml_font(my_style)
    
    # 对比其他样式
    print(f"\n📊 与其他样式对比:")
    compare_with_other_styles(doc, my_style)


def analyze_font_details(style):
    """分析字体详细信息"""
    if hasattr(style, 'font'):
        font = style.font
        
        print(f"  字体名称: {font.name if font.name else '未设置（继承）'}")
        print(f"  字体大小: {font.size if font.size else '未设置（继承）'}")
        print(f"  粗体: {font.bold if font.bold is not None else '未设置（继承）'}")
        print(f"  斜体: {font.italic if font.italic is not None else '未设置（继承）'}")
        print(f"  下划线: {font.underline if font.underline is not None else '未设置（继承）'}")
        
        # 检查颜色
        if font.color and font.color.rgb:
            print(f"  字体颜色: {font.color.rgb}")
        else:
            print(f"  字体颜色: 未设置（继承）")
        
        # 检查中英文字体
        try:
            mixed_fonts = extract_mixed_fonts(font)
            if mixed_fonts:
                print(f"  混合字体设置:")
                for key, value in mixed_fonts.items():
                    if value:
                        print(f"    {key}: {value}")
            else:
                print(f"  混合字体: 未设置（继承）")
        except:
            print(f"  混合字体: 解析失败")
    else:
        print("  ❌ 无字体对象")


def extract_mixed_fonts(font):
    """提取中英文混合字体"""
    mixed_fonts = {}
    try:
        if hasattr(font, '_element'):
            font_element = font._element
            
            # 查找rFonts元素
            rfonts = font_element.find(qn('w:rFonts'))
            if rfonts is not None:
                mixed_fonts["ascii_font"] = rfonts.get(qn('w:ascii'))
                mixed_fonts["hansi_font"] = rfonts.get(qn('w:hAnsi'))
                mixed_fonts["eastasia_font"] = rfonts.get(qn('w:eastAsia'))
                mixed_fonts["cs_font"] = rfonts.get(qn('w:cs'))
        
        return {k: v for k, v in mixed_fonts.items() if v}
    except:
        return {}


def analyze_inheritance(style):
    """分析样式继承关系"""
    try:
        if hasattr(style, 'base_style') and style.base_style:
            base = style.base_style
            print(f"  基础样式: {base.name}")
            print(f"  基础样式类型: {base.type}")
            
            # 分析基础样式的字体
            if hasattr(base, 'font'):
                print(f"  从基础样式继承的字体:")
                font = base.font
                if font.name:
                    print(f"    字体名称: {font.name}")
                if font.size:
                    print(f"    字体大小: {font.size}")
                if font.bold is not None:
                    print(f"    粗体: {font.bold}")
                if font.italic is not None:
                    print(f"    斜体: {font.italic}")
                
                # 检查基础样式的混合字体
                mixed_fonts = extract_mixed_fonts(font)
                if mixed_fonts:
                    print(f"    混合字体:")
                    for key, value in mixed_fonts.items():
                        print(f"      {key}: {value}")
        else:
            print(f"  无基础样式（直接继承默认设置）")
            
            # 查找Normal样式作为默认参考
            for doc_style in style._element.getparent().getparent():
                if hasattr(doc_style, 'styles'):
                    try:
                        normal_style = doc_style.styles['Normal']
                        print(f"  可能继承自Normal样式:")
                        if hasattr(normal_style, 'font') and normal_style.font.name:
                            print(f"    Normal字体: {normal_style.font.name}")
                    except:
                        pass
                    break
    except Exception as e:
        print(f"  继承关系分析失败: {str(e)}")


def analyze_xml_font(style):
    """分析XML中的字体设置"""
    try:
        if hasattr(style, '_element'):
            style_element = style._element
            
            # 查找样式的字体设置
            rpr = style_element.find(qn('w:rPr'))
            if rpr is not None:
                print(f"  找到字体设置元素 (rPr)")
                
                # 查找字体族
                rfonts = rpr.find(qn('w:rFonts'))
                if rfonts is not None:
                    print(f"  字体族设置:")
                    ascii_font = rfonts.get(qn('w:ascii'))
                    hansi_font = rfonts.get(qn('w:hAnsi'))
                    eastasia_font = rfonts.get(qn('w:eastAsia'))
                    cs_font = rfonts.get(qn('w:cs'))
                    
                    if ascii_font:
                        print(f"    ASCII字体（英文）: {ascii_font}")
                    if hansi_font:
                        print(f"    HAnsi字体（英文）: {hansi_font}")
                    if eastasia_font:
                        print(f"    EastAsia字体（中文）: {eastasia_font}")
                    if cs_font:
                        print(f"    ComplexScript字体: {cs_font}")
                    
                    if not any([ascii_font, hansi_font, eastasia_font, cs_font]):
                        print(f"    所有字体设置均为空")
                else:
                    print(f"  未找到字体族设置 (rFonts)")
                
                # 查找字体大小
                sz = rpr.find(qn('w:sz'))
                if sz is not None:
                    size_val = sz.get(qn('w:val'))
                    print(f"  字体大小: {size_val} (半点值)")
                
                # 查找其他字体属性
                if rpr.find(qn('w:b')) is not None:
                    print(f"  粗体: 设置")
                if rpr.find(qn('w:i')) is not None:
                    print(f"  斜体: 设置")
                    
            else:
                print(f"  未找到字体设置元素，样式可能完全继承")
                
                # 查找段落属性中是否有样式引用
                ppr = style_element.find(qn('w:pPr'))
                if ppr is not None:
                    based_on = ppr.find(qn('w:basedOn'))
                    if based_on is not None:
                        based_style = based_on.get(qn('w:val'))
                        print(f"  样式基于: {based_style}")
        else:
            print(f"  无法访问XML元素")
    except Exception as e:
        print(f"  XML分析失败: {str(e)}")


def compare_with_other_styles(doc, my_style):
    """与其他样式对比"""
    print(f"  文档中的其他样式字体设置:")
    
    for style in doc.styles:
        if style.name != "我的样式" and hasattr(style, 'font'):
            font = style.font
            if font.name:  # 只显示有明确字体设置的样式
                print(f"    {style.name}: {font.name}")
                
                # 检查混合字体
                mixed_fonts = extract_mixed_fonts(font)
                if mixed_fonts:
                    font_info = []
                    for key, value in mixed_fonts.items():
                        if value:
                            font_info.append(f"{key}={value}")
                    if font_info:
                        print(f"      混合字体: {', '.join(font_info)}")


def create_font_analysis_report():
    """创建字体分析报告"""
    print(f"\n📄 生成详细分析报告...")
    
    extractor = WordStyleExtractor()
    styles_data = extractor.extract_styles_from_document("../test.docx")
    
    if styles_data and "我的样式" in styles_data.get("styles", {}):
        my_style_data = styles_data["styles"]["我的样式"]
        
        report = {
            "style_name": "我的样式",
            "analysis_summary": {
                "has_font_settings": len(my_style_data.get("font", {})) > 0,
                "has_paragraph_settings": len(my_style_data.get("paragraph", {})) > 0,
                "base_style": my_style_data.get("base_style"),
                "font_details": my_style_data.get("font", {}),
                "paragraph_details": my_style_data.get("paragraph", {})
            },
            "recommendations": []
        }
        
        # 生成建议
        if not my_style_data.get("font", {}):
            report["recommendations"].append("样式未设置字体，将继承父样式或默认字体")
        
        if my_style_data.get("base_style"):
            report["recommendations"].append(f"样式基于: {my_style_data['base_style']['name']}")
        else:
            report["recommendations"].append("样式可能直接继承Normal样式的设置")
        
        # 保存报告
        with open("my_style_font_analysis.json", 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        
        print(f"📋 报告已保存到: my_style_font_analysis.json")
        
        return report
    else:
        print(f"❌ 无法生成报告")
        return None


if __name__ == "__main__":
    analyze_my_style()
    create_font_analysis_report()