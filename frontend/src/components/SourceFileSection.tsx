import React, { useState } from 'react';
import JsonDisplay from './JsonDisplay';
import './SectionComponents.css';

// 类型名称转换
const getTypeName = (type: string): string => {
  const typeNames: { [key: string]: string } = {
    'title': '文档标题',
    'heading1': '一级标题',
    'heading2': '二级标题', 
    'heading3': '三级标题',
    'heading4': '四级标题',
    'paragraph': '正文段落',
    'list': '列表项',
    'quote': '引用',
    'other': '其他'
  };
  return typeNames[type] || type;
};

interface AnalysisSummary {
  total_paragraphs: number;
  average_confidence: number;
  structure_detected: string;
  type_distribution: { [key: string]: number };
}

interface ParseData {
  analysis_summary?: AnalysisSummary;
  [key: string]: any;
}

interface SourceFileSectionProps {
  onComplete: (fileId: string, parseData: ParseData) => void;
}

const SourceFileSection: React.FC<SourceFileSectionProps> = ({ onComplete }) => {
  const [file, setFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [fileId, setFileId] = useState<string>('');
  const [filename, setFilename] = useState<string>('');
  const [parseData, setParseData] = useState<ParseData | null>(null);
  const [parsing, setParsing] = useState(false);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      setFile(selectedFile);
      // 重置状态
      setFileId('');
      setFilename('');
      setParseData(null);
    }
  };

  const handleUpload = async () => {
    if (!file) return;

    setUploading(true);
    try {
      const formData = new FormData();
      formData.append('file', file);

      const response = await fetch('/api/upload/source', {
        method: 'POST',
        body: formData,
      });

      const result = await response.json();
      setFileId(result.file_id);
      setFilename(result.filename);
    } catch (error) {
      console.error('上传失败:', error);
    }
    setUploading(false);
  };

  const handleParse = async () => {
    setParsing(true);
    try {
      const response = await fetch(`/api/parse/source/${fileId}`);
      const result = await response.json();
      setParseData(result);
      onComplete(fileId, result);
    } catch (error) {
      console.error('解析失败:', error);
    }
    setParsing(false);
  };

  return (
    <div className="section-container">
      <h3 className="section-title">源文件解析</h3>
      
      <div className="section-content">
        <div className="upload-area">
          <input
            type="file"
            accept=".docx,.doc"
            onChange={handleFileChange}
            className="file-input"
          />
          {file && (
            <div className="file-info">
              <p>选择的文件: {file.name}</p>
              <button
                onClick={handleUpload}
                disabled={uploading}
                className="action-button"
              >
                {uploading ? '上传中...' : '上传源文件'}
              </button>
            </div>
          )}
        </div>

        {fileId && (
          <div className="parse-area">
            <p className="uploaded-file">已上传: {filename}</p>
            <button
              onClick={handleParse}
              disabled={parsing}
              className="action-button"
            >
              {parsing ? '解析中...' : '解析源文件'}
            </button>
          </div>
        )}

        {parseData && (
          <div className="result-area">
            {/* 显示分析摘要 */}
            {parseData.analysis_summary && (
              <div className="analysis-summary">
                <h4>📊 AI分析摘要</h4>
                <div className="summary-grid">
                  <div className="summary-item">
                    <span className="label">总段落数：</span>
                    <span className="value">{parseData.analysis_summary.total_paragraphs}</span>
                  </div>
                  <div className="summary-item">
                    <span className="label">平均置信度：</span>
                    <span className="value">{parseData.analysis_summary.average_confidence}</span>
                  </div>
                  <div className="summary-item">
                    <span className="label">文档结构：</span>
                    <span className="value">{parseData.analysis_summary.structure_detected}</span>
                  </div>
                </div>
                
                {/* 段落类型分布 */}
                <div className="type-distribution">
                  <h5>段落类型分布：</h5>
                  <div className="type-grid">
                    {Object.entries(parseData.analysis_summary.type_distribution).map(([type, count]: [string, any]) => (
                      <div key={type} className="type-item">
                        <span className="type-name">{getTypeName(type)}</span>
                        <span className="type-count">{count}个</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}
            
            <JsonDisplay 
              data={parseData} 
              title="完整解析结果 (JSON)" 
            />
            <div className="completion-indicator">
              ✅ AI源文件解析完成
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default SourceFileSection;