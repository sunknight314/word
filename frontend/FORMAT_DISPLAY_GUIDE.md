# 格式展示功能使用指南

## 更新内容

现在AI识别格式成功后，不再只显示JSON字符串，而是提供了友好的格式展示功能。

## 使用方法

1. **上传格式要求文档**
   - 选择您的格式要求Word文档
   - 点击"上传格式要求文件"

2. **AI识别格式**
   - 点击"AI生成格式配置"按钮
   - 等待AI分析完成

3. **查看格式展示**
   AI识别成功后，您会看到两个新按钮：
   
   - **📋 查看友好格式展示**：在新标签页中打开格式展示页面
   - **👁️ 在此页面预览**：在当前页面嵌入显示格式（点击可切换显示/隐藏）

4. **查看原始JSON**
   - 原始JSON数据被折叠在"查看原始JSON数据"下
   - 点击即可展开查看

## 格式展示内容

友好的格式展示包括：

- **页面设置**：纸张大小、方向、页边距等
- **样式设置**：各级标题、正文的字体、字号、段落格式等
- **文档结构**：目录、页码等设置

## 技术说明

- 后端API现在返回`view_url`字段，指向格式展示的HTML页面
- 前端组件已更新，自动检测并显示格式展示选项
- 格式数据通过`FormatDisplayHelper`转换为用户友好的描述

## 问题排查

如果您仍然只看到JSON字符串：

1. 确保前端代码已更新（特别是`FormatFileSection.tsx`）
2. 检查浏览器控制台是否有错误
3. 确认后端API返回了`view_url`字段
4. 清除浏览器缓存并刷新页面

## 示例效果

格式展示页面会以卡片形式展示格式要求，包括：
- 清晰的分类和标签
- 中文友好的单位转换（如"12磅（小四）"）
- 视觉化的布局和样式