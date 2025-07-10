# Word文档配置化修改系统

这是一个基于配置文件的Word文档修改系统，可以通过JSON配置文件来批量修改Word文档的样式、结构和内容，而无需修改代码。

## 系统架构

```
document_modify/
├── config_loader.py        # 配置加载器
├── style_modifier.py       # 样式修改器
├── structure_modifier.py   # 结构修改器
├── document_modifier.py    # 主文档修改器
├── modify_config.json      # 示例配置文件
├── example_usage.py        # 使用示例
└── README.md              # 本文档
```

## 功能特性

### 1. 样式修改
- 修改现有样式（字体、大小、颜色、段落格式等）
- 创建新样式
- 样式映射（中文样式名映射到英文）
- 批量应用样式

### 2. 结构修改
- 添加/修改分节符
- 插入/更新目录
- 修改标题层级
- 设置页码格式
- 配置页眉页脚

### 3. 内容修改
- 文本查找和替换（支持正则表达式）
- 插入新内容
- 条件格式化
- 选择性修改

### 4. 页面设置
- 页面大小和方向
- 页边距
- 装订线

## 快速开始

### 1. 基本使用

```python
from document_modifier import DocumentModifier

# 创建修改器
modifier = DocumentModifier("source.docx", "modify_config.json")

# 执行修改
results = modifier.modify_document("output.docx")
```

### 2. 运行示例

```bash
python example_usage.py
```

## 配置文件说明

### 修改模式

```json
{
  "modify_mode": "merge"  // 可选: merge(合并), replace(替换), append(追加), selective(选择性)
}
```

### 样式配置

```json
{
  "styles": {
    "Heading 1": {
      "font": {
        "name": "微软雅黑",
        "size": "16pt",
        "bold": true,
        "color": "#000080"
      },
      "paragraph": {
        "alignment": "left",
        "space_before": "24pt",
        "space_after": "12pt",
        "line_spacing": 1.5
      }
    }
  }
}
```

### 内容修改

```json
{
  "content_modifications": {
    "replacements": [
      {
        "find": "旧文本",
        "replace": "新文本",
        "options": {
          "case_sensitive": false,
          "whole_word": true,
          "use_regex": false
        }
      }
    ]
  }
}
```

### 分节符设置

```json
{
  "section_breaks": {
    "sections": [
      {
        "after_paragraph": 2,
        "type": "new_page"  // 可选: continuous, new_column, new_page, even_page, odd_page
      }
    ]
  }
}
```

## 配置示例

### 示例1：修改标题样式

```json
{
  "styles": {
    "Heading 1": {
      "font": {
        "name": "黑体",
        "size": "18pt",
        "bold": true,
        "color": "#FF0000"
      }
    }
  }
}
```

### 示例2：批量替换文本

```json
{
  "content_modifications": {
    "replacements": [
      {
        "find": "\\{date\\}",
        "replace": "2024年1月",
        "options": {
          "use_regex": true
        }
      }
    ]
  }
}
```

### 示例3：添加页眉页脚

```json
{
  "headers_footers": {
    "header": {
      "text": "文档标题",
      "style": "Header",
      "alignment": "center"
    },
    "footer": {
      "text": "第 {page} 页",
      "style": "Footer",
      "alignment": "center"
    }
  }
}
```

## 高级功能

### 1. 样式映射

处理中文样式名：

```json
{
  "style_mapping": {
    "标题 1": "Heading 1",
    "正文": "Normal"
  }
}
```

### 2. 选择性修改

只修改符合条件的内容：

```json
{
  "selective_modifications": {
    "modifications": [
      {
        "target": "paragraphs",
        "criteria": {
          "style": "Normal",
          "text_contains": "重要"
        },
        "actions": {
          "set_style": "Strong"
        }
      }
    ]
  }
}
```

### 3. 批量操作

```json
{
  "batch_operations": [
    {
      "type": "replace_style",
      "old_style": "Body Text",
      "new_style": "Normal"
    }
  ]
}
```

## 注意事项

1. **备份**: 默认会创建原文档的备份
2. **样式名称**: 使用文档中实际存在的样式名称
3. **编码**: 配置文件使用UTF-8编码
4. **路径**: 使用相对路径或绝对路径都可以

## 错误处理

系统会返回详细的执行结果：

```python
{
  'success': True/False,
  'modifications': {
    'styles': {...},
    'structure': {...},
    'content': {...}
  },
  'errors': [...],
  'warnings': [...]
}
```

## 扩展开发

可以通过继承和扩展各个修改器类来添加新功能：

```python
from style_modifier import StyleModifier

class CustomStyleModifier(StyleModifier):
    def custom_style_operation(self):
        # 自定义操作
        pass
```

## 性能考虑

- 对于大文档，建议分批处理
- 复杂的正则表达式可能影响性能
- 建议先在小文档上测试配置

## 许可证

MIT License