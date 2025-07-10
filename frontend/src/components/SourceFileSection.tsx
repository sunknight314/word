import React, { useState } from 'react';
import DocumentOutlineViewer from './DocumentOutlineViewer';
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
    'abstract_title_cn': '中文摘要标题',
    'abstract_title_en': '英文摘要标题',
    'figure_caption': '图片标题',
    'table_caption': '表格标题',
    'list': '列表项',
    'quote': '引用',
    'other': '其他'
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

interface AnalysisSummary {
  total_paragraphs: number;
  average_confidence: number;
  structure_detected: string;
  type_distribution: { [key: string]: number };
}

interface ParagraphAnalysis {
  paragraph_number: number;
  type: string;
  preview_text?: string;
}

interface ParseData {
  analysis_summary?: AnalysisSummary;
  analysis_result?: ParagraphAnalysis[];
  document_info?: {
    total_paragraphs: number;
    preview_length: number;
    [key: string]: any;
  };
  paragraphs?: Array<{
    paragraph_number: number;
    preview_text: string;
    [key: string]: any;
  }>;
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
            {/* 处理状态 */}
            <div className="parse-status">
              {parseData.success ? (
                <div className="success-indicator">
                  ✅ 文档解析成功
                  {parseData.processing_info?.model_config && (
                    <span className="model-info">
                      (模型: {parseData.processing_info.model_config})
                    </span>
                  )}
                </div>
              ) : (
                <div className="error-indicator">
                  ❌ 文档解析失败: {parseData.error}
                </div>
              )}
            </div>

            {/* 使用DocumentOutlineViewer组件展示目录 */}
            {parseData.analysis_result && parseData.paragraphs && (
              <DocumentOutlineViewer
                analysisResult={parseData.analysis_result}
                paragraphs={parseData.paragraphs}
                documentInfo={parseData.document_info}
              />
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default SourceFileSection;