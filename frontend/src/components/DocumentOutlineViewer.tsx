import React, { useState } from 'react';
import './DocumentOutlineViewer.css';

interface OutlineItem {
  paragraph_number: number;
  type: string;
  text: string;
  level: number;
}

interface DocumentOutlineViewerProps {
  analysisResult: Array<{
    paragraph_number: number;
    type: string;
  }>;
  paragraphs: Array<{
    paragraph_number: number;
    preview_text: string;
  }>;
  documentInfo?: {
    total_paragraphs: number;
    [key: string]: any;
  };
}

const DocumentOutlineViewer: React.FC<DocumentOutlineViewerProps> = ({
  analysisResult,
  paragraphs,
  documentInfo
}) => {
  const [expandedSections, setExpandedSections] = useState<Set<number>>(new Set());
  
  // 类型名称映射
  const getTypeName = (type: string): string => {
    const typeNames: { [key: string]: string } = {
      'Title': '文档标题',
      'Heading1': '一级标题',
      'Heading2': '二级标题',
      'Heading3': '三级标题',
      'Heading4': '四级标题',
      'Normal': '正文',
      'FirstParagraph': '首行缩进段落',
      'AbstractTitleCN': '中文摘要标题',
      'AbstractTitleEN': '英文摘要标题',
      'AbstractContentCN': '中文摘要内容',
      'AbstractContentEN': '英文摘要内容',
      'KeywordsCN': '中文关键词',
      'KeywordsEN': '英文关键词',
      'FigureCaption': '图题',
      'TableCaption': '表题',
      'TOCTitle': '目录标题',
      'TOCItem': '目录项',
      'ReferenceTitle': '参考文献标题',
      'ReferenceItem': '参考文献条目',
      'AcknowledgementTitle': '致谢标题'
    };
    return typeNames[type] || type;
  };
  
  // 判断是否为标题类型
  const isTitleType = (type: string): boolean => {
    const titleTypes = [
      'Title', 'Heading1', 'Heading2', 'Heading3', 'Heading4',
      'AbstractTitleCN', 'AbstractTitleEN', 'TOCTitle', 
      'ReferenceTitle', 'AcknowledgementTitle'
    ];
    return titleTypes.includes(type);
  };
  
  // 获取标题层级
  const getTitleLevel = (type: string): number => {
    const levelMap: { [key: string]: number } = {
      'Title': 0,
      'Heading1': 1,
      'Heading2': 2,
      'Heading3': 3,
      'Heading4': 4,
      'AbstractTitleCN': 1,
      'AbstractTitleEN': 1,
      'TOCTitle': 1,
      'ReferenceTitle': 1,
      'AcknowledgementTitle': 1
    };
    return levelMap[type] || 0;
  };
  
  // 构建目录结构
  const buildOutlineStructure = (): OutlineItem[] => {
    return analysisResult
      .filter(item => isTitleType(item.type))
      .map(item => {
        const paragraph = paragraphs.find(p => p.paragraph_number === item.paragraph_number);
        return {
          paragraph_number: item.paragraph_number,
          type: item.type,
          text: paragraph?.preview_text || '',
          level: getTitleLevel(item.type)
        };
      });
  };
  
  const outlineItems = buildOutlineStructure();
  
  // 统计各类型标题数量
  const getTypeStatistics = () => {
    const stats: { [key: string]: number } = {};
    outlineItems.forEach(item => {
      stats[item.type] = (stats[item.type] || 0) + 1;
    });
    return stats;
  };
  
  const typeStats = getTypeStatistics();
  
  return (
    <div className="outline-viewer">
      <div className="outline-header">
        <h3>📑 文档目录结构</h3>
        <div className="outline-stats">
          <span className="stat-item">共 {outlineItems.length} 个标题</span>
          {documentInfo && (
            <span className="stat-item">总计 {documentInfo.total_paragraphs} 个段落</span>
          )}
        </div>
      </div>
      
      {/* 统计信息 */}
      <div className="type-statistics">
        <h4>标题类型分布</h4>
        <div className="stat-grid">
          {Object.entries(typeStats).map(([type, count]) => (
            <div key={type} className="stat-card">
              <div className="stat-label">{getTypeName(type)}</div>
              <div className="stat-value">{count}</div>
            </div>
          ))}
        </div>
      </div>
      
      {/* 目录内容 */}
      <div className="outline-content">
        <h4>文档结构</h4>
        <div className="outline-tree">
          {outlineItems.map((item, index) => (
            <div
              key={item.paragraph_number}
              className={`outline-node level-${item.level} ${item.type}`}
              style={{ paddingLeft: `${item.level * 24 + 12}px` }}
            >
              <div className="node-content">
                <span className="node-number">{index + 1}.</span>
                <span className="node-text">{item.text}</span>
                <span className="node-info">
                  <span className="node-type">{getTypeName(item.type)}</span>
                  <span className="node-paragraph">第{item.paragraph_number}段</span>
                </span>
              </div>
            </div>
          ))}
        </div>
      </div>
      
      {/* 如果没有找到标题 */}
      {outlineItems.length === 0 && (
        <div className="no-outline">
          <p>未检测到文档标题结构</p>
        </div>
      )}
    </div>
  );
};

export default DocumentOutlineViewer;