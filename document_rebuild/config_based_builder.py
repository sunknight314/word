"""
基于配置的文档构建器 - 主要的重建逻辑（配置驱动版本）
"""

from docx import Document
from content_extractor import load_ai_analysis_result, extract_document_content
from config_loader import FormatConfigLoader
from config_based_styles import create_styles_from_config, apply_page_settings
from config_based_structure import build_structure_from_config
from section_creator import create_toc_section, create_chapter_section, setup_headers, setup_page_numbers_for_all_sections

def create_document_from_config(source_file, output_file, config_path="format_config.json"):
    """基于配置重新创建优化后的文档"""
    
    print("🚀 开始基于配置重新创建优化文档")
    print("=" * 60)
    
    # 第一阶段：配置和内容加载
    print("\n📋 第一阶段：配置和内容加载")
    
    # 1. 加载格式配置
    config_loader = FormatConfigLoader(config_path)
    config = config_loader.load_config()
    if not config:
        print("❌ 无法加载格式配置，退出")
        return False
    
    # 2. 加载AI分析结果
    ai_analysis = load_ai_analysis_result()
    if not ai_analysis:
        print("❌ 无法获取AI分析结果，退出")
        return False
    
    # 3. 提取源文档内容
    content_items = extract_document_content(source_file)
    if not content_items:
        print("❌ 无法提取文档内容，退出")
        return False
    
    # 4. 基于配置构建文档结构
    structure = build_structure_from_config(content_items, ai_analysis, config_loader)
    if not structure['chapters']:
        print("❌ 无法构建文档结构，退出")
        return False
    
    # 第二阶段：新文档创建
    print("\n🏗️ 第二阶段：新文档创建")
    
    # 1. 创建新文档
    doc = Document()
    print("✅ 新文档创建成功")
    
    # 2. 应用页面设置
    page_success = apply_page_settings(doc, config_loader)
    if not page_success:
        print("⚠️ 页面设置应用失败，但继续")
    
    # 3. 基于配置创建样式
    styles = create_styles_from_config(doc, config_loader)
    if not styles:
        print("❌ 创建样式失败，退出")
        return False
    
    # 4. 创建目录节
    toc_success = create_toc_section(doc, structure['toc_section'], styles)
    if not toc_success:
        print("❌ 创建目录节失败")
        return False
    
    # 5. 创建各章节（只创建内容，不设置页码）
    created_sections = []
    for chapter in structure['chapters']:
        section = create_chapter_section(doc, chapter, styles)
        if section:
            created_sections.append(section)
        else:
            print(f"❌ 创建第{chapter['number']}章失败")
            return False
    
    print(f"✅ 成功创建 {len(created_sections)} 个章节")
    
    # 第三阶段：最终优化
    print("\n🎨 第三阶段：最终优化")
    
    # 1. 设置页码
    page_numbers_success = setup_page_numbers_for_all_sections(doc, structure)
    if not page_numbers_success:
        print("⚠️ 页码设置失败，但继续保存文档")
    
    # 2. 设置页眉
    headers_success = setup_headers(doc, structure)
    if not headers_success:
        print("⚠️ 页眉设置失败，但继续保存文档")
    
    # 3. 保存文档
    try:
        doc.save(output_file)
        print(f"💾 文档保存成功: {output_file}")
    except Exception as e:
        print(f"❌ 保存文档失败: {e}")
        return False
    
    # 4. 显示最终统计
    print_config_based_statistics(structure, len(doc.sections), config_loader)
    
    return True

def print_config_based_statistics(structure, total_sections, config_loader):
    """显示基于配置的最终统计信息"""
    
    print("\n📊 最终统计:")
    print(f"  📋 文档标题: {structure['title']['text'] if structure['title'] else '无'}")
    print(f"  📚 章节总数: {len(structure['chapters'])}")
    print(f"  📄 分节总数: {total_sections}")
    
    # 统计各类型段落
    type_stats = {'title': 0, 'heading1': 0, 'heading2': 0, 'heading3': 0, 'paragraph': 0}
    
    if structure['title']:
        type_stats['title'] = 1
    
    for chapter in structure['chapters']:
        for content in chapter['content']:
            content_type = content['type']
            if content_type in type_stats:
                type_stats[content_type] += 1
    
    print(f"  📝 内容统计:")
    print(f"    标题: {type_stats['title']}个")
    print(f"    一级标题: {type_stats['heading1']}个") 
    print(f"    二级标题: {type_stats['heading2']}个")
    print(f"    三级标题: {type_stats['heading3']}个")
    print(f"    正文段落: {type_stats['paragraph']}个")
    
    # 从配置获取页码格式信息
    numbering_config = config_loader.get_page_numbering_config()
    toc_format = numbering_config.get("toc_section", {}).get("format", "upperRoman")
    content_template = numbering_config.get("content_sections", {}).get("template", "第 {page} 页")
    
    print(f"\n🎯 页码格式:")
    if toc_format == "upperRoman":
        print(f"  📋 目录页: 大写罗马数字 (I, II, III...)")
    else:
        print(f"  📋 目录页: 小写罗马数字 (i, ii, iii...)")
    print(f"  📄 正文页: {content_template.replace('{page}', 'X')}")
    
    # 从配置获取页眉信息
    headers_config = config_loader.get_headers_footers_config()
    toc_header = headers_config.get("toc_section", {}).get("odd_header", "目录")
    odd_template = headers_config.get("content_sections", {}).get("odd_header_template", "{chapter_title}")
    even_template = headers_config.get("content_sections", {}).get("even_header_template", "文档标题")
    
    print(f"\n📄 页眉设置:")
    print(f"  📋 目录页: 奇偶页都显示'{toc_header}'")
    print(f"  📚 正文页: 奇数页显示{odd_template}，偶数页显示{even_template}")

def main():
    """主函数"""
    
    print("🎯 Word文档重建测试（配置驱动版本）")
    print("=" * 70)
    
    print("📋 配置驱动模式特点:")
    print("• 所有格式设置从JSON配置文件读取")
    print("• 支持灵活的样式定制")
    print("• 页眉页脚模板化配置")
    print("• 页码格式可配置")
    print("• 页面设置可配置")
    print("• 文档结构可配置")
    print()
    
    # 文件路径
    source_file = "../test_files/test_document.docx"
    output_file = "test_document_config_based.docx"
    config_file = "format_config.json"
    
    # 执行重建
    success = create_document_from_config(source_file, output_file, config_file)
    
    if success:
        print(f"\n🎉 配置驱动重建完成!")
        print(f"📝 源文件: {source_file}")
        print(f"📝 重建后: {output_file}")
        print(f"⚙️ 配置文件: {config_file}")
        print(f"💡 在Word中打开查看效果，右键更新目录")
    else:
        print(f"\n❌ 配置驱动重建失败!")

if __name__ == "__main__":
    main()