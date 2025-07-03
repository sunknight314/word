from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

# 创建新文档
doc = Document()

current_section = doc.sections[-1]  # last section in document
current_section.start_type

new_section = doc.add_section(WD_SECTION.ODD_PAGE)
new_section.start_type
print(new_section.iter_inner_content())
# 保存文档
doc.save("分节页码测试.docx")
print("文档已生成：包含奇数页分节符和统一页码")