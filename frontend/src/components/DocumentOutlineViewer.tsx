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
      'title': '文档标题',
      'heading1': '一级标题',
      'heading2': '二级标题',
      'heading3': '三级标题',
      'heading4': '四级标题',
      'abstract_title_cn': '中文摘要',
      'abstract_title_en': '英文摘要',
      'figure_caption': '图片标题',
      'table_caption': '表格标题'
    };
    return typeNames[type] || type;
  };
  
  // 判断是否为标题类型
  const isTitleType = (type: string): boolean => {
    const titleTypes = [
      'title', 'heading1', 'heading2', 'heading3', 'heading4',
      'abstract_title_cn', 'abstract_title_en'
    ];
    return titleTypes.includes(type);
  };
  
  // 获取标题层级
  const getTitleLevel = (type: string): number => {
    const levelMap: { [key: string]: number } = {
      'title': 0,
      'heading1': 1,
      'heading2': 2,
      'heading3': 3,
      'heading4': 4,
      'abstract_title_cn': 1,
      'abstract_title_en': 1
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