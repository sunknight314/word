# Word文档样式提取工具

这个工具用于提取Word文档中的所有样式信息，包括字体设置、段落格式等详细信息。

## 文件说明

- `style_extractor.py` - 主要的样式提取程序
- `README.md` - 使用说明文档

## 功能特点

✅ **完整样式提取** - 提取文档中所有样式（段落样式、字符样式、表格样式等）  
✅ **详细格式信息** - 包含字体、段落格式、对齐方式、缩进等所有设置  
✅ **中英文字体支持** - 提取中英文混合字体设置  
✅ **大纲级别识别** - 识别标题的大纲级别  
✅ **JSON格式输出** - 结构化数据便于后续处理  
✅ **样式统计分析** - 按类型统计样式数量  

## 安装依赖

```bash
pip install python-docx
```

## 使用方法

### 1. 基本使用

```python
from style_extractor import WordStyleExtractor

# 创建提取器
extractor = WordStyleExtractor()

# 提取样式
styles_data = extractor.extract_styles_from_document("your_document.docx")

# 打印摘要
extractor.print_summary(styles_data)

# 保存为JSON
extractor.save_to_json(styles_data, "styles_output.json")
```

### 2. 直接运行

修改 `style_extractor.py` 中的文件路径，然后运行：

```bash
python style_extractor.py
```

### 3. 批量处理

```python
import os
from pathlib import Path

extractor = WordStyleExtractor()

# 处理文件夹中的所有docx文件
folder_path = "path/to/your/documents"
for docx_file in Path(folder_path).glob("*.docx"):
    styles_data = extractor.extract_styles_from_document(str(docx_file))
    output_name = f"styles_{docx_file.stem}.json"
    extractor.save_to_json(styles_data, output_name)
```

## 输出格式

提取的样式信息以JSON格式保存，包含以下结构：

```json
{
  "document_info": {
    "file_name": "文档名称.docx",
    "total_styles": 样式总数,
    "style_type_count": {
      "段落样式": 数量,
      "字符样式": 数量,
      "表格样式": 数量
    }
  },
  "styles": {
    "样式名称": {
      "name": "样式名称",
      "type": "样式类型",
      "builtin": true/false,
      "font": {
        "name": "字体名称",
        "size": "字号pt",
        "bold": true/false,
        "italic": true/false,
        "ascii_font": "英文字体",
        "eastasia_font": "中文字体"
      },
      "paragraph": {
        "alignment": "对齐方式",
        "line_spacing": "行距pt",
        "space_before": "段前距pt",
        "space_after": "段后距pt",
        "first_line_indent": "首行缩进pt",
        "left_indent": "左缩进pt",
        "right_indent": "右缩进pt"
      },
      "outline_level": 大纲级别
    }
  }
}
```

## 支持的样式属性

### 字体属性
- 字体名称（中英文分别）
- 字号
- 粗体/斜体
- 下划线/删除线
- 字体颜色
- 小型大写字母
- 全大写字母
- 阴影/轮廓效果

### 段落属性  
- 对齐方式（左对齐、居中、右对齐、两端对齐）
- 行距（固定值、倍数行距等）
- 段前距/段后距
- 首行缩进/左缩进/右缩进
- 孤行控制
- 段落分页设置
- 大纲级别

### 样式属性
- 样式类型（段落/字符/表格/列表）
- 是否为内置样式
- 是否隐藏
- 优先级
- 快速样式设置

## 注意事项

1. **文件路径** - 确保Word文档路径正确且文件存在
2. **权限** - 确保有读取文档和写入JSON文件的权限
3. **文档格式** - 仅支持.docx格式（不支持.doc）
4. **中文支持** - 输出的JSON文件使用UTF-8编码，支持中文

## 示例输出

运行后会在控制台看到类似输出：

```
📖 正在提取文档样式: 测试文档.docx
📊 文档中共有 85 个样式
✅ 样式提取完成

============================================================
📋 样式提取摘要
============================================================
📄 文档名称: 测试文档.docx
📊 样式总数: 85

📈 样式类型统计:
  段落样式: 42个
  字符样式: 28个
  表格样式: 15个

📝 样式列表 (前10个):
   1. Normal (段落样式) [内置:✅]
   2. Heading 1 (段落样式) [内置:✅]
   3. Heading 2 (段落样式) [内置:✅]
   4. Title (段落样式) [内置:✅]
   5. CustomTitle (段落样式) [内置:❌]
   ... 还有 80 个样式

💾 样式信息已保存到: styles_测试文档.json
```

## 扩展功能

如需要其他功能，可以扩展 `WordStyleExtractor` 类：

- 添加表格样式的详细提取
- 增加列表样式的编号设置
- 提取页面设置信息
- 添加样式之间的继承关系分析