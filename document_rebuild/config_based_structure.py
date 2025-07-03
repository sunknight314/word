"""
基于配置的文档结构构建器 - 从JSON配置构建文档结构
"""

from config_loader import FormatConfigLoader

def build_structure_from_config(content_items, ai_analysis, config_loader):
    """基于配置构建文档结构"""
    
    print("🏗️ 基于配置构建文档结构...")
    
    # 获取各种配置
    doc_info = config_loader.get_document_info()
    toc_config = config_loader.get_toc_config()
    headers_config = config_loader.get_headers_footers_config()
    structure_config = config_loader.get_document_structure_config()
    
    structure = {
        'title': None,
        'toc_section': build_toc_section_config(toc_config, headers_config),
        'chapters': []
    }
    
    current_chapter = None
    chapter_number = 0
    
    for item in content_items:
        para_num = item['paragraph_number']
        para_type = ai_analysis.get(para_num, 'paragraph')
        
        # 处理文档标题
        if para_type == 'title' and not structure['title']:
            title_text = item['text']
            # 如果配置中有指定标题，使用配置的；否则使用提取的
            if doc_info.get('title'):
                title_text = doc_info['title']
            
            structure['title'] = {
                'text': title_text,
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
            current_chapter = build_chapter_config(
                chapter_number, item, headers_config, structure_config
            )
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

def build_toc_section_config(toc_config, headers_config):
    """构建目录节配置"""
    toc_section = {
        'title': toc_config.get('title', '目录'),
        'page_format': 'roman',
        'page_start': toc_config.get('page_start', 1),
        'headers': {}
    }
    
    # 从headers_config获取目录页眉设置
    if headers_config.get('toc_section'):
        toc_headers = headers_config['toc_section']
        toc_section['headers'] = {
            'odd': toc_headers.get('odd_header', '目录'),
            'even': toc_headers.get('even_header', '目录')
        }
    else:
        toc_section['headers'] = {'odd': '目录', 'even': '目录'}
    
    return toc_section

def build_chapter_config(chapter_number, item, headers_config, structure_config):
    """构建章节配置"""
    
    # 获取页眉模板
    content_headers = headers_config.get('content_sections', {})
    odd_template = content_headers.get('odd_header_template', '{chapter_title}')
    even_template = content_headers.get('even_header_template', 'Word文档格式优化项目')
    
    # 替换模板变量
    odd_header = odd_template.replace('{chapter_title}', item['text'])
    odd_header = odd_header.replace('{chapter_number}', str(chapter_number))
    
    even_header = even_template.replace('{chapter_title}', item['text'])
    even_header = even_header.replace('{chapter_number}', str(chapter_number))
    
    chapter = {
        'number': chapter_number,
        'title': item['text'],
        'page_break': structure_config.get('chapter_start_page', 'odd_page'),
        'page_format': 'decimal',
        'page_restart': True if chapter_number == 1 else False,
        'page_start': 1 if chapter_number == 1 else None,
        'headers': {
            'odd': odd_header,
            'even': even_header
        },
        'content': [{
            'text': item['text'],
            'type': 'heading1',
            'paragraph_number': item['paragraph_number']
        }]
    }
    
    return chapter

def get_section_break_type(config_loader, break_type):
    """获取分节符类型"""
    breaks_config = config_loader.get_section_breaks_config()
    
    if break_type == "between_toc_and_content":
        return breaks_config.get("between_toc_and_content", "odd_page")
    elif break_type == "between_chapters":
        return breaks_config.get("between_chapters", "odd_page")
    else:
        return breaks_config.get("between_sections", "continuous")

def get_page_numbering_format(config_loader, section_type):
    """获取页码格式"""
    numbering_config = config_loader.get_page_numbering_config()
    
    if section_type == "toc":
        toc_config = numbering_config.get("toc_section", {})
        return {
            'format': toc_config.get('format', 'upperRoman'),
            'start': toc_config.get('start', 1),
            'restart': toc_config.get('restart', True),
            'template': toc_config.get('template', '{page}')
        }
    else:
        content_config = numbering_config.get("content_sections", {})
        return {
            'format': content_config.get('format', 'decimal'),
            'start': content_config.get('start', 1),
            'restart_first': content_config.get('restart_first_chapter', True),
            'continue_others': content_config.get('continue_others', True),
            'template': content_config.get('template', '第 {page} 页')
        }