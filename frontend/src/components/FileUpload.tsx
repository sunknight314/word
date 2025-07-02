import React, { useState } from 'react';
import './FileUpload.css';

const FileUpload: React.FC = () => {
  const [originalFile, setOriginalFile] = useState<File | null>(null);
  const [formatFile, setFormatFile] = useState<File | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    console.log('文件上传功能待实现');
  };

  return (
    <div className="file-upload">
      <form onSubmit={handleSubmit}>
        <div className="upload-section">
          <label htmlFor="original-file">原始文档：</label>
          <input
            id="original-file"
            type="file"
            accept=".docx,.doc"
            onChange={(e) => setOriginalFile(e.target.files?.[0] || null)}
          />
        </div>
        
        <div className="upload-section">
          <label htmlFor="format-file">格式要求文档：</label>
          <input
            id="format-file"
            type="file"
            accept=".docx,.doc"
            onChange={(e) => setFormatFile(e.target.files?.[0] || null)}
          />
        </div>
        
        <button 
          type="submit" 
          disabled={!originalFile || !formatFile}
          className="upload-btn"
        >
          开始处理
        </button>
      </form>
    </div>
  );
};

export default FileUpload;