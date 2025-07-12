"""
Word样式提取工具使用示例
演示各种使用场景和高级功能
"""

import os
import json
from pathlib import Path
from style_extractor import WordStyleExtractor


def example_basic_usage():
    """示例1: 基本使用方法"""
    print("🔍 示例1: 基本样式提取")
    print("-" * 40)
    
    extractor = WordStyleExtractor()
    
    # 指定要分析的文档
    doc_path = "../test.docx"  # 请修改为实际的文档路径
    
    if Path(doc_path).exists():
        # 提取样式
        styles_data = extractor.extract_styles_from_document(doc_path)
        
        if styles_data:
            # 显示摘要
            extractor.print_summary(styles_data)
            
            # 保存结果
            extractor.save_to_json(styles_data, "basic_styles.json")
        else:
            print("❌ 样式提取失败")
    else:
        print(f"⚠️  文件不存在: {doc_path}")


def example_batch_processing():
    """示例2: 批量处理多个文档"""
    print("\n🔍 示例2: 批量处理文档")
    print("-" * 40)
    
    extractor = WordStyleExtractor()
    
    # 要处理的文档列表
    documents = [
        "../test.docx",
        "../1、西安电子科技大学研究生学位论文模板（2015年版）-2022.11修订.docx",
        # 可以添加更多文档
    ]
    
    results = {}
    
    for doc_path in documents:
        if Path(doc_path).exists():
            print(f"\n📄 处理文档: {Path(doc_path).name}")
            
            styles_data = extractor.extract_styles_from_document(doc_path)
            if styles_data:
                doc_name = Path(doc_path).stem
                results[doc_name] = styles_data
                
                # 为每个文档保存单独的JSON文件
                output_file = f"batch_styles_{doc_name}.json"
                extractor.save_to_json(styles_data, output_file)
        else:
            print(f"⚠️  跳过不存在的文件: {doc_path}")
    
    # 保存批量处理的汇总结果
    if results:
        with open("batch_processing_summary.json", 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"\n💾 批量处理汇总已保存到: batch_processing_summary.json")


def example_style_analysis():
    """示例3: 样式详细分析"""
    print("\n🔍 示例3: 样式详细分析")
    print("-" * 40)
    
    extractor = WordStyleExtractor()
    doc_path = "../test.docx"  # 请修改为实际的文档路径
    
    if not Path(doc_path).exists():
        print(f"⚠️  文件不存在: {doc_path}")
        return
    
    styles_data = extractor.extract_styles_from_document(doc_path)
    if not styles_data:
        print("❌ 样式提取失败")
        return
    
    styles = styles_data.get("styles", {})
    
    print(f"\n📊 详细分析报告:")
    
    # 1. 按类型分组分析
    style_groups = {}
    for name, info in styles.items():
        style_type = info.get("type", "未知")
        if style_type not in style_groups:
            style_groups[style_type] = []
        style_groups[style_type].append((name, info))
    
    for style_type, style_list in style_groups.items():
        print(f"\n📋 {style_type} ({len(style_list)}个):")
        for i, (name, info) in enumerate(style_list[:5]):  # 只显示前5个
            builtin = "内置" if info.get("builtin", False) else "自定义"
            print(f"  {i+1}. {name} ({builtin})")
        if len(style_list) > 5:
            print(f"  ... 还有 {len(style_list) - 5} 个")
    
    # 2. 分析标题样式
    print(f"\n📝 标题样式分析:")
    heading_styles = []
    for name, info in styles.items():
        if "heading" in name.lower() or "标题" in name or info.get("outline_level") is not None:
            heading_styles.append((name, info))
    
    for name, info in sorted(heading_styles, key=lambda x: x[1].get("outline_level", 999))[:6]:
        outline = info.get("outline_level", "无")
        font_size = "未知"
        if "font" in info and "size" in info["font"]:
            font_size = info["font"]["size"]
        print(f"  {name}: 大纲级别={outline}, 字号={font_size}")
    
    # 3. 分析字体使用情况
    print(f"\n🔤 字体使用统计:")
    font_usage = {}
    for name, info in styles.items():
        if "font" in info:
            font_info = info["font"]
            # 统计中文字体
            chinese_font = font_info.get("eastasia_font", font_info.get("name", "未知"))
            if chinese_font != "未知":
                font_usage[chinese_font] = font_usage.get(chinese_font, 0) + 1
    
    for font, count in sorted(font_usage.items(), key=lambda x: x[1], reverse=True)[:5]:
        print(f"  {font}: {count}个样式使用")
    
    # 4. 分析自定义样式
    custom_styles = [(name, info) for name, info in styles.items() 
                    if not info.get("builtin", True)]
    
    print(f"\n🎨 自定义样式 ({len(custom_styles)}个):")
    for name, info in custom_styles[:8]:  # 只显示前8个
        style_type = info.get("type", "未知")
        print(f"  {name} ({style_type})")
    if len(custom_styles) > 8:
        print(f"  ... 还有 {len(custom_styles) - 8} 个自定义样式")


def example_compare_documents():
    """示例4: 比较两个文档的样式差异"""
    print("\n🔍 示例4: 文档样式比较")
    print("-" * 40)
    
    extractor = WordStyleExtractor()
    
    # 要比较的两个文档
    doc1_path = "../test.docx"
    doc2_path = "../1、西安电子科技大学研究生学位论文模板（2015年版）-2022.11修订.docx"
    
    if not (Path(doc1_path).exists() and Path(doc2_path).exists()):
        print("⚠️  请确保两个比较文档都存在")
        return
    
    print(f"📄 文档1: {Path(doc1_path).name}")
    styles1 = extractor.extract_styles_from_document(doc1_path)
    
    print(f"📄 文档2: {Path(doc2_path).name}")
    styles2 = extractor.extract_styles_from_document(doc2_path)
    
    if not (styles1 and styles2):
        print("❌ 样式提取失败")
        return
    
    styles1_names = set(styles1["styles"].keys())
    styles2_names = set(styles2["styles"].keys())
    
    # 比较分析
    print(f"\n📊 样式比较结果:")
    print(f"文档1样式数量: {len(styles1_names)}")
    print(f"文档2样式数量: {len(styles2_names)}")
    
    common_styles = styles1_names & styles2_names
    unique_to_doc1 = styles1_names - styles2_names
    unique_to_doc2 = styles2_names - styles1_names
    
    print(f"\n🤝 共同样式 ({len(common_styles)}个):")
    for style in sorted(list(common_styles))[:8]:
        print(f"  {style}")
    if len(common_styles) > 8:
        print(f"  ... 还有 {len(common_styles) - 8} 个")
    
    print(f"\n📄 文档1独有样式 ({len(unique_to_doc1)}个):")
    for style in sorted(list(unique_to_doc1))[:5]:
        print(f"  {style}")
    if len(unique_to_doc1) > 5:
        print(f"  ... 还有 {len(unique_to_doc1) - 5} 个")
    
    print(f"\n📄 文档2独有样式 ({len(unique_to_doc2)}个):")
    for style in sorted(list(unique_to_doc2))[:5]:
        print(f"  {style}")
    if len(unique_to_doc2) > 5:
        print(f"  ... 还有 {len(unique_to_doc2) - 5} 个")
    
    # 保存比较结果
    comparison_result = {
        "document1": {
            "name": Path(doc1_path).name,
            "total_styles": len(styles1_names),
            "unique_styles": list(unique_to_doc1)
        },
        "document2": {
            "name": Path(doc2_path).name,
            "total_styles": len(styles2_names),
            "unique_styles": list(unique_to_doc2)
        },
        "common_styles": list(common_styles),
        "comparison_summary": {
            "total_common": len(common_styles),
            "doc1_unique": len(unique_to_doc1),
            "doc2_unique": len(unique_to_doc2)
        }
    }
    
    with open("document_styles_comparison.json", 'w', encoding='utf-8') as f:
        json.dump(comparison_result, f, ensure_ascii=False, indent=2)
    print(f"\n💾 比较结果已保存到: document_styles_comparison.json")


def example_export_specific_styles():
    """示例5: 导出特定类型的样式"""
    print("\n🔍 示例5: 导出特定样式")
    print("-" * 40)
    
    extractor = WordStyleExtractor()
    doc_path = "../test.docx"  # 请修改为实际的文档路径
    
    if not Path(doc_path).exists():
        print(f"⚠️  文件不存在: {doc_path}")
        return
    
    styles_data = extractor.extract_styles_from_document(doc_path)
    if not styles_data:
        print("❌ 样式提取失败")
        return
    
    styles = styles_data["styles"]
    
    # 1. 导出所有标题样式
    heading_styles = {}
    for name, info in styles.items():
        if ("heading" in name.lower() or "标题" in name or 
            info.get("outline_level") is not None):
            heading_styles[name] = info
    
    with open("heading_styles_only.json", 'w', encoding='utf-8') as f:
        json.dump(heading_styles, f, ensure_ascii=False, indent=2)
    print(f"📝 标题样式已导出到: heading_styles_only.json ({len(heading_styles)}个)")
    
    # 2. 导出自定义样式
    custom_styles = {name: info for name, info in styles.items() 
                    if not info.get("builtin", True)}
    
    with open("custom_styles_only.json", 'w', encoding='utf-8') as f:
        json.dump(custom_styles, f, ensure_ascii=False, indent=2)
    print(f"🎨 自定义样式已导出到: custom_styles_only.json ({len(custom_styles)}个)")
    
    # 3. 导出段落样式
    paragraph_styles = {name: info for name, info in styles.items() 
                       if info.get("type") == "段落样式"}
    
    with open("paragraph_styles_only.json", 'w', encoding='utf-8') as f:
        json.dump(paragraph_styles, f, ensure_ascii=False, indent=2)
    print(f"📄 段落样式已导出到: paragraph_styles_only.json ({len(paragraph_styles)}个)")


def main():
    """运行所有示例"""
    print("🚀 Word样式提取工具 - 使用示例")
    print("=" * 60)
    
    # 运行各种示例
    example_basic_usage()
    example_batch_processing()
    example_style_analysis()
    example_compare_documents()
    example_export_specific_styles()
    
    print("\n" + "=" * 60)
    print("✅ 所有示例运行完成！")
    print("📁 生成的文件:")
    
    # 列出生成的文件
    generated_files = [
        "basic_styles.json",
        "batch_processing_summary.json",
        "document_styles_comparison.json",
        "heading_styles_only.json",
        "custom_styles_only.json",
        "paragraph_styles_only.json"
    ]
    
    for filename in generated_files:
        if Path(filename).exists():
            size = Path(filename).stat().st_size
            print(f"  📄 {filename} ({size} 字节)")


if __name__ == "__main__":
    main()