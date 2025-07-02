import React, { useState } from 'react';
import JsonDisplay from './JsonDisplay';
import './SectionComponents.css';

interface FormatFileSectionProps {
  onComplete: (fileId: string, parseData: any) => void;
}

const FormatFileSection: React.FC<FormatFileSectionProps> = ({ onComplete }) => {
  const [file, setFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [fileId, setFileId] = useState<string>('');
  const [filename, setFilename] = useState<string>('');
  const [parseData, setParseData] = useState<any>(null);
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
            <JsonDisplay 
              data={parseData} 
              title="格式要求解析结果" 
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