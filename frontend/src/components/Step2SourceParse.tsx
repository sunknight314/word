import React, { useState, useEffect } from 'react';
import JsonDisplay from './JsonDisplay';
import './StepComponents.css';

interface Step2Props {
  fileId: string;
  filename: string;
  onParseSuccess: (parseData: any) => void;
}

const Step2SourceParse: React.FC<Step2Props> = ({ fileId, filename, onParseSuccess }) => {
  const [parseData, setParseData] = useState<any>(null);
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
                <span className="success-indicator">✅ 段落分析成功</span>
              ) : (
                <span className="error-indicator">❌ 段落分析失败: {parseData.error}</span>
              )}
              {parseData.processing_info?.model_config && (
                <span className="model-info">
                  (模型: {parseData.processing_info.model_config})
                </span>
              )}
            </div>
            
            {/* 显示段落分析JSON */}
            {parseData.paragraph_analysis_json && (
              <JsonDisplay 
                data={parseData.paragraph_analysis_json} 
                title="📋 段落类型分析结果 (JSON)" 
              />
            )}
            
            {/* 显示完整解析结果 */}
            <JsonDisplay 
              data={parseData} 
              title="📄 完整源文件解析结果" 
            />
          </div>
        )}
      </div>
    </div>
  );
};

export default Step2SourceParse;