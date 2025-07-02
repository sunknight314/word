import React, { useState } from 'react';
import SourceFileSection from '../components/SourceFileSection';
import FormatFileSection from '../components/FormatFileSection';
import ProcessSection from '../components/ProcessSection';
import './UploadPage.css';

const UploadPage: React.FC = () => {
  const [sourceFileId, setSourceFileId] = useState<string>('');
  const [sourceParseData, setSourceParseData] = useState<any>(null);
  
  const [formatFileId, setFormatFileId] = useState<string>('');
  const [formatParseData, setFormatParseData] = useState<any>(null);

  const handleSourceComplete = (fileId: string, parseData: any) => {
    setSourceFileId(fileId);
    setSourceParseData(parseData);
  };

  const handleFormatComplete = (fileId: string, parseData: any) => {
    setFormatFileId(fileId);
    setFormatParseData(parseData);
  };

  const canProcess = sourceFileId && formatFileId && sourceParseData && formatParseData;

  return (
    <div className="upload-page">
      <h2>Word文档格式优化</h2>
      <p>三个操作部分可以独立进行，完成前两部分后即可进行格式转换：</p>
      
      <div className="sections-container">
        {/* 源文件解析部分 */}
        <SourceFileSection onComplete={handleSourceComplete} />
        
        {/* 格式要求文件解析部分 */}
        <FormatFileSection onComplete={handleFormatComplete} />
        
        {/* 格式转换部分 */}
        <ProcessSection 
          sourceFileId={sourceFileId}
          formatFileId={formatFileId}
          canProcess={canProcess}
        />
      </div>
    </div>
  );
};

export default UploadPage;