"""
文档构建器 - 主要的重建逻辑
"""

from docx import Document
from content_extractor import load_ai_analysis_result, extract_document_content, build_document_structure
from styles_manager import create_document_styles
from section_creator import create_toc_section, create_chapter_section, setup_headers, setup_page_numbers_for_all_sections

def create_optimized_document(source_file, output_file):
    """重新创建优化后的文档"""
    
    print("🚀 开始重新创建优化文档")
    print("=" * 60)
    
    # 第一阶段：内容提取与分析
    print("\n📖 第一阶段：内容提取与分析")
    
    # 1. 加载AI分析结果
    ai_analysis = load_ai_analysis_result()
    if not ai_analysis:
        print("❌ 无法获取AI分析结果，退出")
        return False
    
    # 2. 提取源文档内容
    content_items = extract_document_content(source_file)
    if not content_items:
        print("❌ 无法提取文档内容，退出")
        return False
    
    # 3. 构建文档结构
    structure = build_document_structure(content_items, ai_analysis)
    if not structure['chapters']:
        print("❌ 无法构建文档结构，退出")
        return False
    
    # 第二阶段：新文档创建
    print("\n🏗️ 第二阶段：新文档创建")
    
    # 1. 创建新文档
    doc = Document()
    print("✅ 新文档创建成功")
    
    # 2. 创建样式
    styles = create_document_styles(doc)
    if not styles:
        print("❌ 创建样式失败，退出")
        return False
    
    # 3. 创建目录节
    toc_success = create_toc_section(doc, structure['toc_section'], styles)
    if not toc_success:
        print("❌ 创建目录节失败")
        return False
    
    # 4. 创建各章节（只创建内容，不设置页码）
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
    
    # 3. 显示最终统计
    print_final_statistics(structure, len(doc.sections))
    
    return True

def print_final_statistics(structure, total_sections):
    """显示最终统计信息"""
    
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
    
    print(f"\n🎯 页码格式:")
    print(f"  📋 目录页: 罗马数字 (I, II, III...)")
    print(f"  📄 正文页: 阿拉伯数字 (第1页, 第2页...)")
    
    print(f"\n📄 页眉设置:")
    print(f"  📋 目录页: 奇偶页都显示'目录'")
    print(f"  📚 正文页: 奇数页显示章标题，偶数页显示'Word文档格式优化项目'")

def main():
    """主函数"""
    
    print("🎯 Word文档重建测试")
    print("=" * 70)
    
    print("📋 新建模式特点:")
    print("• 从零创建文档，避免python-docx API限制")
    print("• 目录直接在开头创建，无需移动")
    print("• 分节符在创建时设置，无需后插入")
    print("• 页码从创建时就是正确格式")
    print("• 最大化使用高级API，最小化XML操作")
    print("• 中英文混合字体完美支持")
    print()
    
    # 文件路径
    source_file = "../test_files/test_document.docx"
    output_file = "test_document_rebuilt.docx"
    
    # 执行重建
    success = create_optimized_document(source_file, output_file)
    
    if success:
        print(f"\n🎉 重建完成!")
        print(f"📝 源文件: {source_file}")
        print(f"📝 重建后: {output_file}")
        print(f"💡 在Word中打开查看效果，右键更新目录")
    else:
        print(f"\n❌ 重建失败!")

if __name__ == "__main__":
    main()