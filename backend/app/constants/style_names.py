"""
标准样式名称定义
统一整个系统中使用的样式名称
"""

class StyleNames:
    """标准样式名称常量"""
    
    # 基础样式
    TITLE = "Title"                          # 标题
    HEADING1 = "Heading1"                    # 一级标题
    HEADING2 = "Heading2"                    # 二级标题
    HEADING3 = "Heading3"                    # 三级标题
    HEADING4 = "Heading4"                    # 四级标题
    NORMAL = "Normal"                        # 正文
    FIRST_PARAGRAPH = "FirstParagraph"       # 首行缩进段落
    
    # 摘要相关
    ABSTRACT_TITLE_CN = "AbstractTitleCN"    # 中文摘要标题
    ABSTRACT_TITLE_EN = "AbstractTitleEN"    # 英文摘要标题
    ABSTRACT_CONTENT_CN = "AbstractContentCN" # 中文摘要内容
    ABSTRACT_CONTENT_EN = "AbstractContentEN" # 英文摘要内容
    
    # 关键词
    KEYWORDS_CN = "KeywordsCN"               # 中文关键词
    KEYWORDS_EN = "KeywordsEN"               # 英文关键词
    
    # 目录相关
    TOC_TITLE = "TOCTitle"                   # 目录标题
    TOC_ITEM = "TOCItem"                     # 目录项
    
    # 图表标题
    FIGURE_CAPTION = "FigureCaption"         # 图题
    TABLE_CAPTION = "TableCaption"           # 表题
    
    # 页眉页脚
    HEADER = "Header"                        # 页眉
    FOOTER = "Footer"                        # 页脚
    
    # 特殊内容
    FOOTNOTE = "Footnote"                    # 脚注
    QUOTE = "Quote"                          # 引用
    CODE = "Code"                            # 代码
    
    # 列表
    LIST = "List"                            # 列表
    BULLET_LIST = "BulletList"               # 项目符号列表
    NUMBERED_LIST = "NumberedList"           # 编号列表
    
    # 学术论文特有
    CHAPTER_TITLE = "ChapterTitle"           # 章标题
    SECTION_TITLE = "SectionTitle"           # 节标题
    SUBSECTION_TITLE = "SubsectionTitle"     # 小节标题
    THESIS_TITLE = "ThesisTitle"             # 论文标题
    AUTHOR_INFO = "AuthorInfo"               # 作者信息
    AUTHOR_INFO_CN = "AuthorInfoCN"          # 中文作者信息
    AUTHOR_INFO_EN = "AuthorInfoEN"          # 英文作者信息
    
    # 参考文献
    REFERENCE = "Reference"                  # 参考文献
    REFERENCE_TITLE = "ReferenceTitle"       # 参考文献标题
    REFERENCE_ITEM = "ReferenceItem"         # 参考文献条目
    
    # 附录
    APPENDIX = "Appendix"                    # 附录
    APPENDIX_TITLE = "AppendixTitle"         # 附录标题
    
    # 致谢
    ACKNOWLEDGEMENT = "Acknowledgement"      # 致谢
    ACKNOWLEDGEMENT_TITLE = "AcknowledgementTitle" # 致谢标题


# 样式名称到中文描述的映射
STYLE_DESCRIPTIONS = {
    StyleNames.TITLE: "标题",
    StyleNames.HEADING1: "一级标题",
    StyleNames.HEADING2: "二级标题",
    StyleNames.HEADING3: "三级标题",
    StyleNames.HEADING4: "四级标题",
    StyleNames.NORMAL: "正文",
    StyleNames.FIRST_PARAGRAPH: "首行缩进段落",
    
    StyleNames.ABSTRACT_TITLE_CN: "中文摘要标题",
    StyleNames.ABSTRACT_TITLE_EN: "英文摘要标题",
    StyleNames.ABSTRACT_CONTENT_CN: "中文摘要内容",
    StyleNames.ABSTRACT_CONTENT_EN: "英文摘要内容",
    
    StyleNames.KEYWORDS_CN: "中文关键词",
    StyleNames.KEYWORDS_EN: "英文关键词",
    
    StyleNames.TOC_TITLE: "目录标题",
    StyleNames.TOC_ITEM: "目录项",
    
    StyleNames.FIGURE_CAPTION: "图题",
    StyleNames.TABLE_CAPTION: "表题",
    
    StyleNames.HEADER: "页眉",
    StyleNames.FOOTER: "页脚",
    
    StyleNames.FOOTNOTE: "脚注",
    StyleNames.QUOTE: "引用",
    StyleNames.CODE: "代码",
    
    StyleNames.LIST: "列表",
    StyleNames.BULLET_LIST: "项目符号列表",
    StyleNames.NUMBERED_LIST: "编号列表",
    
    StyleNames.CHAPTER_TITLE: "章标题",
    StyleNames.SECTION_TITLE: "节标题",
    StyleNames.SUBSECTION_TITLE: "小节标题",
    StyleNames.THESIS_TITLE: "论文标题",
    StyleNames.AUTHOR_INFO: "作者信息",
    StyleNames.AUTHOR_INFO_CN: "中文作者信息",
    StyleNames.AUTHOR_INFO_EN: "英文作者信息",
    
    StyleNames.REFERENCE: "参考文献",
    StyleNames.REFERENCE_TITLE: "参考文献标题",
    StyleNames.REFERENCE_ITEM: "参考文献条目",
    
    StyleNames.APPENDIX: "附录",
    StyleNames.APPENDIX_TITLE: "附录标题",
    
    StyleNames.ACKNOWLEDGEMENT: "致谢",
    StyleNames.ACKNOWLEDGEMENT_TITLE: "致谢标题",
}


# 样式分类
STYLE_CATEGORIES = {
    "基础样式": [
        StyleNames.TITLE,
        StyleNames.HEADING1,
        StyleNames.HEADING2,
        StyleNames.HEADING3,
        StyleNames.HEADING4,
        StyleNames.NORMAL,
        StyleNames.FIRST_PARAGRAPH,
    ],
    "摘要相关": [
        StyleNames.ABSTRACT_TITLE_CN,
        StyleNames.ABSTRACT_TITLE_EN,
        StyleNames.ABSTRACT_CONTENT_CN,
        StyleNames.ABSTRACT_CONTENT_EN,
        StyleNames.KEYWORDS_CN,
        StyleNames.KEYWORDS_EN,
    ],
    "目录": [
        StyleNames.TOC_TITLE,
        StyleNames.TOC_ITEM,
    ],
    "图表": [
        StyleNames.FIGURE_CAPTION,
        StyleNames.TABLE_CAPTION,
    ],
    "页面元素": [
        StyleNames.HEADER,
        StyleNames.FOOTER,
        StyleNames.FOOTNOTE,
    ],
    "特殊格式": [
        StyleNames.QUOTE,
        StyleNames.CODE,
        StyleNames.LIST,
        StyleNames.BULLET_LIST,
        StyleNames.NUMBERED_LIST,
    ],
    "学术论文": [
        StyleNames.CHAPTER_TITLE,
        StyleNames.SECTION_TITLE,
        StyleNames.SUBSECTION_TITLE,
        StyleNames.THESIS_TITLE,
        StyleNames.AUTHOR_INFO,
        StyleNames.AUTHOR_INFO_CN,
        StyleNames.AUTHOR_INFO_EN,
        StyleNames.REFERENCE,
        StyleNames.REFERENCE_TITLE,
        StyleNames.REFERENCE_ITEM,
        StyleNames.APPENDIX,
        StyleNames.APPENDIX_TITLE,
        StyleNames.ACKNOWLEDGEMENT,
        StyleNames.ACKNOWLEDGEMENT_TITLE,
    ],
}


def get_style_description(style_name: str) -> str:
    """获取样式的中文描述"""
    return STYLE_DESCRIPTIONS.get(style_name, style_name)


def get_all_style_names() -> list:
    """获取所有样式名称列表"""
    return list(STYLE_DESCRIPTIONS.keys())


def is_valid_style_name(style_name: str) -> bool:
    """检查样式名称是否有效"""
    return style_name in STYLE_DESCRIPTIONS