import React, { useState, useEffect } from 'react';
import DocumentOutlineViewer from './DocumentOutlineViewer';
import './StepComponents.css';

interface ParseData {
  success: boolean;
  error?: string;
  processing_info?: {
    model_config?: string;
    [key: string]: any;
  };
  analysis_result?: Array<{
    paragraph_number: number;
    type: string;
  }>;
  paragraphs?: Array<{
    paragraph_number: number;
    preview_text: string;
  }>;
  document_info?: {
    total_paragraphs: number;
    [key: string]: any;
  };
  [key: string]: any;
}

interface Step2Props {
  fileId: string;
  filename: string;
  onParseSuccess: (parseData: any) => void;
}

const Step2SourceParse: React.FC<Step2Props> = ({ fileId, filename, onParseSuccess }) => {
  const [parseData, setParseData] = useState<ParseData | null>(null);
  const [loading, setLoading] = useState(false);

  const handleParse = async () => {
    setLoading(true);
    try {
      const response = await fetch(`/api/parse/source/${fileId}`);
      const result = await response.json();
      setParseData(result);
      onParseSuccess(result);
    } catch (error) {
      console.error('解析失败:', error);
    }
    setLoading(false);
  };

  return (
    <div className="step-container">
      <h3>步骤 2: 解析源文件</h3>
      <div className="step-content">
        <p className="file-info">文件: {filename}</p>
        <button
          onClick={handleParse}
          disabled={loading}
          className="step-button"
        >
          {loading ? '解析中...' : '解析源文件'}
        </button>
        
        {parseData && (
          <div>
            {/* 显示处理状态 */}
            <div className="result-status">
              {parseData.success ? (
                <span className="success-indicator">✅ 文档解析成功</span>
              ) : (
                <span className="error-indicator">❌ 文档解析失败: {parseData.error}</span>
              )}
              {parseData.processing_info?.model_config && (
                <span className="model-info">
                  (模型: {parseData.processing_info.model_config})
                </span>
              )}
            </div>
            
            {/* 使用DocumentOutlineViewer展示文档目录 */}
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

export default Step2SourceParse;