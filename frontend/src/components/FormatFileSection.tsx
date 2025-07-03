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
  onComplete?: (fileId: string, parseData: FormatData) => void;
}

const FormatFileSection: React.FC<FormatFileSectionProps> = ({ onComplete }) => {
  const [file, setFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [fileId, setFileId] = useState<string>('');
  const [filename, setFilename] = useState<string>('');
  const [aiGenerating, setAiGenerating] = useState(false);
  const [formatConfig, setFormatConfig] = useState<any>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      setFile(selectedFile);
      // 重置状态
      setFileId('');
      setFilename('');
      setFormatConfig(null);
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

  const handleAIGenerate = async () => {
    setAiGenerating(true);
    try {
      const response = await fetch(`/api/generate/format-config/${fileId}`, {
        method: 'POST',
      });
      const result = await response.json();
      setFormatConfig(result);
    } catch (error) {
      console.error('AI生成格式配置失败:', error);
    }
    setAiGenerating(false);
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
            <div className="button-group">
              <button
                onClick={handleAIGenerate}
                disabled={aiGenerating}
                className="action-button ai-button"
              >
                {aiGenerating ? 'AI生成中...' : 'AI生成格式配置'}
              </button>
            </div>
          </div>
        )}

        {/* AI生成的格式配置显示 */}
        {formatConfig && (
          <div className="ai-config-area">
            <h4>🤖 AI生成的格式配置</h4>
            {formatConfig.success ? (
              <div className="ai-success">
                <div className="config-status">
                  ✅ AI生成成功
                  {formatConfig.ai_info?.model && (
                    <span className="model-info">
                      (模型: {formatConfig.ai_info.model.model || 'DeepSeek-V2.5'})
                    </span>
                  )}
                </div>
                <JsonDisplay 
                  data={formatConfig.format_config} 
                  title="生成的format_config.json" 
                />
                {formatConfig.document_info && (
                  <div className="document-summary">
                    <p><strong>文档信息:</strong> {formatConfig.document_info.total_paragraphs} 段落, {formatConfig.document_info.total_length} 字符</p>
                  </div>
                )}
              </div>
            ) : (
              <div className="ai-error">
                <div className="error-status">
                  ❌ AI生成失败: {formatConfig.error}
                  {formatConfig.fallback && <span className="fallback-note">(已使用默认配置)</span>}
                </div>
                {formatConfig.format_config && (
                  <JsonDisplay 
                    data={formatConfig.format_config} 
                    title="默认格式配置" 
                  />
                )}
              </div>
            )}
          </div>
        )}

      </div>
    </div>
  );
};

export default FormatFileSection;