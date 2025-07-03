import React, { useState } from 'react';
import JsonDisplay from './JsonDisplay';
import './SectionComponents.css';

// 格式分析摘要接口
interface FormatSummary {
  total_styles: number;
  main_fonts: string[];
  font_size_range: {
    min: number;
    max: number;
    count: number;
  };
  paper_info: {
    size: string;
    orientation: string;
    margins_uniform: boolean;
  };
  format_features: string[];
  format_complexity: string;
}

interface FormatData {
  analysis_summary?: FormatSummary;
  recommendations?: string[];
  [key: string]: any;
}

interface FormatFileSectionProps {
  onComplete: (fileId: string, parseData: FormatData) => void;
}

const FormatFileSection: React.FC<FormatFileSectionProps> = ({ onComplete }) => {
  const [file, setFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [fileId, setFileId] = useState<string>('');
  const [filename, setFilename] = useState<string>('');
  const [parseData, setParseData] = useState<FormatData | null>(null);
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

      const response = await fetch('/api/upload/format', {
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
      const response = await fetch(`/api/parse/format/${fileId}`);
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
      <h3 className="section-title">格式要求文件解析</h3>
      
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
                {uploading ? '上传中...' : '上传格式要求文件'}
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
              {parsing ? '解析中...' : '解析格式要求文件'}
            </button>
          </div>
        )}

        {parseData && (
          <div className="result-area">
            {/* 显示格式分析摘要 */}
            {parseData.analysis_summary && (
              <div className="analysis-summary">
                <h4>📋 格式分析摘要</h4>
                <div className="summary-grid">
                  <div className="summary-item">
                    <span className="label">样式总数：</span>
                    <span className="value">{parseData.analysis_summary.total_styles}</span>
                  </div>
                  <div className="summary-item">
                    <span className="label">格式复杂度：</span>
                    <span className="value">{parseData.analysis_summary.format_complexity}</span>
                  </div>
                  <div className="summary-item">
                    <span className="label">纸张信息：</span>
                    <span className="value">{parseData.analysis_summary.paper_info.size} - {parseData.analysis_summary.paper_info.orientation}</span>
                  </div>
                </div>
                
                {/* 主要字体 */}
                <div className="font-info">
                  <h5>主要字体：</h5>
                  <div className="font-list">
                    {parseData.analysis_summary.main_fonts.map((font: string, index: number) => (
                      <span key={index} className="font-tag">{font}</span>
                    ))}
                  </div>
                </div>
                
                {/* 格式特征 */}
                <div className="features-info">
                  <h5>格式特征：</h5>
                  <div className="features-list">
                    {parseData.analysis_summary.format_features.map((feature: string, index: number) => (
                      <span key={index} className="feature-tag">{feature}</span>
                    ))}
                  </div>
                </div>
              </div>
            )}
            
            {/* 显示优化建议 */}
            {parseData.recommendations && parseData.recommendations.length > 0 && (
              <div className="recommendations">
                <h4>💡 格式优化建议</h4>
                <ul className="recommendation-list">
                  {parseData.recommendations.map((rec: string, index: number) => (
                    <li key={index} className="recommendation-item">{rec}</li>
                  ))}
                </ul>
              </div>
            )}
            
            <JsonDisplay 
              data={parseData} 
              title="完整格式解析结果 (JSON)" 
            />
            <div className="completion-indicator">
              ✅ 格式要求文件解析完成
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default FormatFileSection;