# 基于配置的Word文档格式化系统使用说明

## 概述

这是一个基于JSON配置文件的Word文档格式化系统，允许用户通过修改配置文件来自定义所有格式设置，而无需修改代码。

## 核心文件

### 1. 配置文件
- `format_config.json` - 主配置文件，包含所有格式设置

### 2. 核心模块
- `config_loader.py` - 配置加载器，负责读取和解析JSON配置
- `config_based_styles.py` - 基于配置的样式管理器
- `config_based_structure.py` - 基于配置的文档结构构建器
- `config_based_builder.py` - 主要的重建逻辑（配置驱动版本）

## 使用方法

### 基本使用
```bash
python3 config_based_builder.py
```

### 自定义配置文件
```python
from config_based_builder import create_document_from_config

success = create_document_from_config(
    source_file="input.docx",
    output_file="output.docx", 
    config_path="my_custom_config.json"
)
```

## 配置文件结构详解

### 1. 文档信息 (`document_info`)
```json
{
  "document_info": {
    "title": "文档标题",
    "author": "作者姓名",
    "description": "文档描述"
  }
}
```

### 2. 页面设置 (`page_settings`)
```json
{
  "page_settings": {
    "margins": {
      "top": "2.54cm",      // 支持: pt, cm, in
      "bottom": "2.54cm",
      "left": "3.17cm", 
      "right": "3.17cm"
    },
    "orientation": "portrait",  // portrait | landscape
    "size": "A4"               // 目前仅支持描述
  }
}
```

### 3. 样式设置 (`styles`)
每种内容类型都可以自定义样式：

```json
{
  "styles": {
    "title": {
      "name": "CustomTitle",           // 样式名称
      "font": {
        "chinese": "黑体",            // 中文字体
        "english": "Times New Roman", // 英文字体
        "size": "22pt",              // 字体大小
        "bold": true,                // 是否加粗
        "italic": false              // 是否斜体
      },
      "paragraph": {
        "alignment": "center",        // left|center|right|justify
        "line_spacing": "20pt",       // 行距
        "space_before": "0pt",        // 段前距
        "space_after": "24pt",        // 段后距
        "first_line_indent": "0pt",   // 首行缩进
        "left_indent": "0pt",         // 左缩进（整段左移）
        "right_indent": "0pt",        // 右缩进（整段右移）
        "hanging_indent": "0pt"       // 悬挂缩进（除首行外的所有行缩进）
      },
      "outline_level": null           // 大纲级别（null或0-8）
    }
  }
}
```

支持的内容类型：
- `title` - 文档标题
- `heading1` - 一级标题
- `heading2` - 二级标题 
- `heading3` - 三级标题
- `paragraph` - 正文段落

### 4. 目录设置 (`toc_settings`)
```json
{
  "toc_settings": {
    "title": "目录",              // 目录标题
    "levels": "1-3",             // 包含的标题级别
    "hyperlinks": true,          // 是否包含超链接
    "tab_leader": "dots",        // 制表符样式
    "page_format": "roman",      // 页码格式
    "page_start": 1,             // 起始页码
    "headers": {
      "odd": "目录",
      "even": "目录"
    }
  }
}
```

### 5. 页码设置 (`page_numbering`)
```json
{
  "page_numbering": {
    "toc_section": {
      "format": "upperRoman",     // upperRoman|lowerRoman|decimal
      "start": 1,
      "restart": true,
      "template": "{page}"        // 页码模板
    },
    "content_sections": {
      "format": "decimal",
      "start": 1,
      "restart_first_chapter": true,
      "continue_others": true,
      "template": "第 {page} 页"
    }
  }
}
```

### 6. 页眉页脚设置 (`headers_footers`)
```json
{
  "headers_footers": {
    "odd_even_different": true,         // 奇偶页不同
    "first_page_different": false,     // 首页不同
    "toc_section": {
      "odd_header": "目录",
      "even_header": "目录"
    },
    "content_sections": {
      "odd_header_template": "{chapter_title}",    // 支持变量
      "even_header_template": "Word文档格式优化项目"
    }
  }
}
```

支持的模板变量：
- `{chapter_title}` - 章节标题
- `{chapter_number}` - 章节编号

### 7. 分节符设置 (`section_breaks`)
```json
{
  "section_breaks": {
    "between_toc_and_content": "odd_page",   // 目录和正文间
    "between_chapters": "odd_page",          // 章节间
    "between_sections": "continuous"         // 其他分节间
  }
}
```

分节符类型：
- `continuous` - 连续
- `new_column` - 新列
- `new_page` - 新页
- `even_page` - 偶数页
- `odd_page` - 奇数页

### 8. 文档结构设置 (`document_structure`)
```json
{
  "document_structure": {
    "auto_generate_toc": true,        // 自动生成目录
    "chapter_start_page": "odd",      // 章节起始页类型
    "chapter_numbering": "arabic",    // 章节编号格式
    "section_numbering": "decimal"    // 分节编号格式
  }
}
```

## 自定义示例

### 修改字体和大小
```json
{
  "styles": {
    "heading1": {
      "font": {
        "chinese": "微软雅黑",
        "english": "Arial",
        "size": "18pt",
        "bold": true
      }
    }
  }
}
```

### 设置段落缩进
```json
{
  "styles": {
    "paragraph": {
      "paragraph": {
        "first_line_indent": "24pt",    // 首行缩进2字符
        "left_indent": "0pt",           // 整段左缩进
        "right_indent": "0pt",          // 整段右缩进
        "hanging_indent": "0pt"         // 悬挂缩进
      }
    },
    "heading2": {
      "paragraph": {
        "left_indent": "12pt",          // 二级标题整体缩进
        "hanging_indent": "0pt"
      }
    }
  }
}
```

### 创建引用样式（悬挂缩进示例）
```json
{
  "styles": {
    "quote": {
      "name": "CustomQuote",
      "font": {
        "chinese": "楷体",
        "english": "Times New Roman",
        "size": "11pt",
        "italic": true
      },
      "paragraph": {
        "alignment": "left",
        "left_indent": "36pt",          // 整段左缩进3字符
        "right_indent": "36pt",         // 整段右缩进3字符
        "hanging_indent": "12pt",       // 悬挂缩进1字符
        "space_before": "6pt",
        "space_after": "6pt"
      }
    }
  }
}
```

### 修改页眉模板
```json
{
  "headers_footers": {
    "content_sections": {
      "odd_header_template": "第{chapter_number}章 - {chapter_title}",
      "even_header_template": "我的文档标题"
    }
  }
}
```

### 修改页码格式
```json
{
  "page_numbering": {
    "content_sections": {
      "template": "- {page} -"
    }
  }
}
```

## 优势

1. **配置驱动** - 无需修改代码，只需修改配置文件
2. **灵活定制** - 支持字体、段落、页码、页眉等全方位定制
3. **模板化** - 页眉支持变量模板，便于批量处理
4. **向前兼容** - 与原有system保持兼容，可并行使用
5. **易于维护** - 配置与代码分离，便于维护和分发

## 注意事项

1. 配置文件必须是有效的JSON格式
2. 长度单位支持：`pt`（磅）、`cm`（厘米）、`in`（英寸）
3. 修改配置后无需重启，直接运行即可生效
4. 建议在修改配置前备份原始配置文件

### 缩进设置说明

1. **首行缩进** (`first_line_indent`)：只影响段落的第一行
2. **左缩进** (`left_indent`)：整个段落从左边距向右移动
3. **右缩进** (`right_indent`)：整个段落从右边距向左移动  
4. **悬挂缩进** (`hanging_indent`)：除第一行外的所有行都缩进

### 缩进组合效果

- `first_line_indent: "24pt"` + `left_indent: "0pt"` = 首行缩进2字符
- `first_line_indent: "0pt"` + `left_indent: "24pt"` = 整段缩进2字符
- `hanging_indent: "12pt"` + `left_indent: "24pt"` = 悬挂缩进效果

### 常用缩进值参考

- 1个中文字符 ≈ 12pt
- 2个中文字符 ≈ 24pt  
- 3个中文字符 ≈ 36pt