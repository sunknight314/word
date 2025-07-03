"""
内容提取器 - 从源文档中提取纯内容
"""

import json
from docx import Document

def load_ai_analysis_result():
    """加载AI分析结果"""
    
    result_file = "../backend/ai_analysis_results/ai_analysis_success_20250702_184058.json"
    
    try:
        with open(result_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 提取分析结果
        analysis_results = data["analysis_result"]["result"]["analysis_result"]
        
        # 创建段落类型映射
        paragraph_types = {}
        for item in analysis_results:
            para_num = item["paragraph_number"]
            para_type = item["type"]
            paragraph_types[para_num] = para_type
        
        print(f"✅ 成功加载AI分析结果: {len(paragraph_types)}个段落")
        return paragraph_types
        
    except Exception as e:
        print(f"❌ 加载AI分析结果失败: {e}")
        return {}

def extract_document_content(source_file):
    """提取源文档的纯内容"""
    
    print(f"📄 提取文档内容: {source_file}")
    
    try:
        doc = Document(source_file)
        
        content_items = []
        for i, para in enumerate(doc.paragraphs, 1):
            if para.text.strip():  # 跳过空段落
                content_items.append({
                    'paragraph_number': i,
                    'text': para.text.strip(),
                    'original_style': para.style.name,
                    'has_runs': len(para.runs) > 0
                })
        
        print(f"✅ 提取完成: {len(content_items)}个有效段落")
        return content_items
        
    except Exception as e:
        print(f"❌ 提取内容失败: {e}")
        return []

def build_document_structure(content_items, ai_analysis):
    """根据AI分析构建文档结构"""
    
    print(f"🏗️ 构建文档结构...")
    
    structure = {
        'title': None,
        'toc_section': {
            'title': '目录',
            'page_format': 'roman',
            'page_start': 1,
            'headers': {'odd': '目录', 'even': '目录'}
        },
        'chapters': []
    }
    
    current_chapter = None
    chapter_number = 0
    
    for item in content_items:
        para_num = item['paragraph_number']
        para_type = ai_analysis.get(para_num, 'paragraph')
        
        # 处理文档标题
        if para_type == 'title' and not structure['title']:
            structure['title'] = {
                'text': item['text'],
                'type': 'title'
            }
            continue
        
        # 处理一级标题 - 新章节
        if para_type == 'heading1':
            # 保存前一章节
            if current_chapter:
                structure['chapters'].append(current_chapter)
            
            # 开始新章节
            chapter_number += 1
            current_chapter = {
                'number': chapter_number,
                'title': item['text'],
                'page_break': 'odd_page',
                'page_format': 'decimal',
                'page_restart': True if chapter_number == 1 else False,
                'page_start': 1 if chapter_number == 1 else None,
                'headers': {
                    'odd': item['text'],  # 章标题作为奇数页页眉
                    'even': "Word文档格式优化项目"
                },
                'content': [{
                    'text': item['text'],
                    'type': para_type,
                    'paragraph_number': para_num
                }]
            }
        else:
            # 添加到当前章节
            if current_chapter:
                current_chapter['content'].append({
                    'text': item['text'],
                    'type': para_type,
                    'paragraph_number': para_num
                })
    
    # 添加最后一章
    if current_chapter:
        structure['chapters'].append(current_chapter)
    
    print(f"✅ 结构构建完成:")
    print(f"  📋 标题: {structure['title']['text'] if structure['title'] else '无'}")
    print(f"  📚 章节数: {len(structure['chapters'])}")
    
    for i, chapter in enumerate(structure['chapters'], 1):
        print(f"    第{i}章: {chapter['title']} ({len(chapter['content'])}段)")
    
    return structure