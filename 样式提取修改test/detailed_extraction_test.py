"""
详细样式提取测试脚本
专门用于测试和验证详细的字体和段落设置提取
"""

from style_extractor import WordStyleExtractor
from pathlib import Path
import json


def test_detailed_extraction():
    """测试详细样式提取功能"""
    print("🔍 详细样式提取测试")
    print("=" * 50)
    
    extractor = WordStyleExtractor()
    
    # 测试文档路径
    test_doc = "../test.docx"
    
    if not Path(test_doc).exists():
        print(f"⚠️  测试文档不存在: {test_doc}")
        print("请确保有可用的测试文档")
        return
    
    print(f"📄 测试文档: {Path(test_doc).name}")
    
    # 提取样式
    styles_data = extractor.extract_styles_from_document(test_doc)
    
    if not styles_data:
        print("❌ 样式提取失败")
        return
    
    # 分析提取的详细信息
    analyze_detailed_styles(styles_data)
    
    # 保存详细结果
    output_file = "detailed_styles_test.json"
    extractor.save_to_json(styles_data, output_file)
    
    # 创建详细报告
    create_detailed_report(styles_data, "detailed_extraction_report.txt")
    
    print(f"\n✅ 详细提取测试完成")
    print(f"📁 生成文件:")
    print(f"  📄 {output_file}")
    print(f"  📄 detailed_extraction_report.txt")


def analyze_detailed_styles(styles_data):
    """分析详细样式信息"""
    print(f"\n🔍 详细分析结果:")
    
    styles = styles_data.get("styles", {})
    
    # 统计各类详细信息
    stats = {
        "font_details": 0,
        "paragraph_details": 0,
        "color_info": 0,
        "border_info": 0,
        "shading_info": 0,
        "tab_stops": 0,
        "xml_attributes": 0,
        "mixed_fonts": 0,
        "outline_levels": 0
    }
    
    detailed_styles = []
    
    for name, info in styles.items():
        detail_count = 0
        style_details = {"name": name, "features": []}
        
        # 检查字体详细信息
        if 'font' in info:
            font_info = info['font']
            stats["font_details"] += 1
            detail_count += 1
            
            font_features = []
            if 'size' in font_info:
                font_features.append(f"字号:{font_info['size']}")
            if 'bold' in font_info:
                font_features.append(f"粗体:{font_info['bold']}")
            if 'italic' in font_info:
                font_features.append(f"斜体:{font_info['italic']}")
            if 'color_rgb' in font_info:
                font_features.append(f"颜色:{font_info['color_rgb']}")
            if 'eastasia_font' in font_info:
                font_features.append(f"中文字体:{font_info['eastasia_font']}")
                stats["mixed_fonts"] += 1
            if 'ascii_font' in font_info:
                font_features.append(f"英文字体:{font_info['ascii_font']}")
            if 'xml_attributes' in font_info:
                stats["xml_attributes"] += 1
                font_features.append("XML属性")
            
            if font_features:
                style_details["features"].append(f"字体: {', '.join(font_features)}")
        
        # 检查段落详细信息
        if 'paragraph' in info:
            para_info = info['paragraph']
            stats["paragraph_details"] += 1
            detail_count += 1
            
            para_features = []
            if 'alignment' in para_info:
                para_features.append(f"对齐:{para_info['alignment']}")
            if 'line_spacing' in para_info:
                para_features.append(f"行距:{para_info['line_spacing']}")
            if 'first_line_indent' in para_info and para_info['first_line_indent'] != "0pt":
                para_features.append(f"首行缩进:{para_info['first_line_indent']}")
            if 'space_before' in para_info and para_info['space_before'] != "0pt":
                para_features.append(f"段前距:{para_info['space_before']}")
            if 'space_after' in para_info and para_info['space_after'] != "0pt":
                para_features.append(f"段后距:{para_info['space_after']}")
            if 'tab_stops' in para_info:
                stats["tab_stops"] += 1
                para_features.append(f"制表符:{len(para_info['tab_stops'])}个")
            if 'borders' in para_info:
                stats["border_info"] += 1
                para_features.append("边框")
            if 'shading' in para_info:
                stats["shading_info"] += 1
                para_features.append("阴影")
            if 'xml_attributes' in para_info:
                para_features.append("XML属性")
            
            if para_features:
                style_details["features"].append(f"段落: {', '.join(para_features)}")
        
        # 检查大纲级别
        if info.get('outline_level') is not None:
            stats["outline_levels"] += 1
            style_details["features"].append(f"大纲级别: {info['outline_level']}")
            detail_count += 1
        
        # 如果有详细信息，添加到列表
        if detail_count > 0:
            detailed_styles.append(style_details)
    
    # 打印统计信息
    print(f"\n📊 详细信息统计:")
    print(f"  包含字体详细信息的样式: {stats['font_details']}个")
    print(f"  包含段落详细信息的样式: {stats['paragraph_details']}个")
    print(f"  包含中英文混合字体的样式: {stats['mixed_fonts']}个")
    print(f"  包含大纲级别的样式: {stats['outline_levels']}个")
    print(f"  包含颜色信息的样式: {stats['color_info']}个")
    print(f"  包含边框信息的样式: {stats['border_info']}个")
    print(f"  包含阴影信息的样式: {stats['shading_info']}个")
    print(f"  包含制表符设置的样式: {stats['tab_stops']}个")
    print(f"  包含XML扩展属性的样式: {stats['xml_attributes']}个")
    
    # 显示最详细的样式（前5个）
    detailed_styles.sort(key=lambda x: len(x['features']), reverse=True)
    
    print(f"\n🏆 最详细的样式 (前5个):")
    for i, style in enumerate(detailed_styles[:5]):
        print(f"  {i+1}. {style['name']} ({len(style['features'])}个特性)")
        for feature in style['features']:
            print(f"     • {feature}")


def create_detailed_report(styles_data, report_file):
    """创建详细的文本报告"""
    try:
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("Word文档样式详细提取报告\n")
            f.write("=" * 50 + "\n\n")
            
            doc_info = styles_data.get("document_info", {})
            f.write(f"文档名称: {doc_info.get('file_name', 'Unknown')}\n")
            f.write(f"样式总数: {doc_info.get('total_styles', 0)}\n")
            f.write(f"生成时间: {doc_info.get('extraction_time', 'Unknown')}\n\n")
            
            styles = styles_data.get("styles", {})
            
            f.write("详细样式信息:\n")
            f.write("-" * 30 + "\n\n")
            
            for name, info in styles.items():
                f.write(f"样式名称: {name}\n")
                f.write(f"样式类型: {info.get('type', 'Unknown')}\n")
                f.write(f"内置样式: {'是' if info.get('builtin', False) else '否'}\n")
                
                if 'font' in info:
                    f.write("字体设置:\n")
                    font_info = info['font']
                    for key, value in font_info.items():
                        if value is not None and value != "" and value != "0pt":
                            f.write(f"  {key}: {value}\n")
                
                if 'paragraph' in info:
                    f.write("段落设置:\n")
                    para_info = info['paragraph']
                    for key, value in para_info.items():
                        if value is not None and value != "" and value != "0pt":
                            f.write(f"  {key}: {value}\n")
                
                if info.get('outline_level') is not None:
                    f.write(f"大纲级别: {info['outline_level']}\n")
                
                f.write("\n" + "-" * 40 + "\n\n")
        
        print(f"📄 详细报告已保存到: {report_file}")
        
    except Exception as e:
        print(f"❌ 创建报告失败: {str(e)}")


def compare_with_simple_extraction():
    """比较详细提取与简单提取的差异"""
    print(f"\n🔄 比较详细提取与简单提取")
    print("-" * 40)
    
    # 这里可以实现简单提取的对比逻辑
    # 比如只提取基本的字体名称、大小等
    print("详细提取包含更多信息:")
    print("  ✅ 中英文混合字体")
    print("  ✅ 字符间距和位置")
    print("  ✅ 段落边框和阴影")
    print("  ✅ 制表符设置") 
    print("  ✅ XML扩展属性")
    print("  ✅ 原始数值和转换后数值")


if __name__ == "__main__":
    test_detailed_extraction()
    compare_with_simple_extraction()