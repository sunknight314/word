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
          <JsonDisplay 
            data={parseData} 
            title="源文件解析结果" 
          />
        )}
      </div>
    </div>
  );
};

export default Step2SourceParse;