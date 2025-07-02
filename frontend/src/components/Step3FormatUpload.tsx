import React, { useState } from 'react';
import './StepComponents.css';

interface Step3Props {
  onUploadSuccess: (fileId: string, filename: string) => void;
}

const Step3FormatUpload: React.FC<Step3Props> = ({ onUploadSuccess }) => {
  const [file, setFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      setFile(selectedFile);
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
      onUploadSuccess(result.file_id, result.filename);
    } catch (error) {
      console.error('上传失败:', error);
    }
    setUploading(false);
  };

  return (
    <div className="step-container">
      <h3>步骤 3: 上传格式要求文件</h3>
      <div className="step-content">
        <input
          type="file"
          accept=".docx,.doc"
          onChange={handleFileChange}
          className="file-input"
        />
        {file && (
          <p className="file-info">选择的文件: {file.name}</p>
        )}
        <button
          onClick={handleUpload}
          disabled={!file || uploading}
          className="step-button"
        >
          {uploading ? '上传中...' : '上传格式要求文件'}
        </button>
      </div>
    </div>
  );
};

export default Step3FormatUpload;