from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def insert_odd_page_section_break(doc, paragraph_index):
    """
    在指定段落后插入真正的奇数页分节符
    
    参数:
        doc: Document对象
        paragraph_index: 目标段落索引(0开始)
    """
    # 获取目标段落
    target_paragraph = doc.paragraphs[paragraph_index]
    
    # 创建新段落用于分节符
    new_paragraph = OxmlElement('w:p')
    
    # 创建段落属性
    p_pr = OxmlElement('w:pPr')
    new_paragraph.append(p_pr)
    
    # 创建分节属性
    sect_pr = OxmlElement('w:sectPr')
    p_pr.append(sect_pr)
    
    # 设置分节类型为奇数页
    type_element = OxmlElement('w:type')
    type_element.set(qn('w:val'), 'oddPage')
    sect_pr.append(type_element)
    
    # 添加页面尺寸
    pg_sz = OxmlElement('w:pgSz')
    pg_sz.set(qn('w:w'), '12240')
    pg_sz.set(qn('w:h'), '15840')
    sect_pr.append(pg_sz)
    
    # 添加页面边距
    pg_mar = OxmlElement('w:pgMar')
    pg_mar.set(qn('w:top'), '1440')
    pg_mar.set(qn('w:right'), '1440')
    pg_mar.set(qn('w:bottom'), '1440')
    pg_mar.set(qn('w:left'), '1440')
    pg_mar.set(qn('w:header'), '720')
    pg_mar.set(qn('w:footer'), '720')
    pg_mar.set(qn('w:gutter'), '0')
    sect_pr.append(pg_mar)
    
    # 在目标段落后插入分节符段落
    target_paragraph._p.addnext(new_paragraph)
    
    # 返回新创建的段落以便后续操作
    return new_paragraph

# 使用示例
doc = Document("test_document.docx")
insert_odd_page_section_break(doc, 2)  # 在第三段后插入
doc.save("modified_document.docx")