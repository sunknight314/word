import React, { useState } from 'react';
import './SectionComponents.css';

interface ProcessSectionProps {
  sourceFileId: string;
  formatFileId: string;
  canProcess: boolean;
}

const ProcessSection: React.FC<ProcessSectionProps> = ({ 
  sourceFileId, 
  formatFileId, 
  canProcess 
}) => {
  const [processing, setProcessing] = useState(false);
  const [taskId, setTaskId] = useState<string | null>(null);
  const [status, setStatus] = useState<string>('');
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const handleProcess = async () => {
    if (!canProcess) return;

    setProcessing(true);
    try {
      // 开始处理
      const processResponse = await fetch('/api/process', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          source_id: sourceFileId,
          format_id: formatFileId,
        }),
      });

      const processResult = await processResponse.json();
      setTaskId(processResult.task_id);
      setStatus(processResult.status);

      // 轮询状态
      const pollStatus = async () => {
        const statusResponse = await fetch(`/api/status/${processResult.task_id}`);
        const statusResult = await statusResponse.json();
        
        setStatus(statusResult.status);
        
        if (statusResult.status === 'completed') {
          setDownloadUrl(statusResult.download_url);
          setProcessing(false);
        } else if (statusResult.status === 'failed') {
          setProcessing(false);
        } else {
          setTimeout(pollStatus, 2000);
        }
      };

      setTimeout(pollStatus, 1000);
    } catch (error) {
      console.error('处理失败:', error);
      setProcessing(false);
    }
  };

  const handleDownload = () => {
    if (downloadUrl) {
      window.open(downloadUrl, '_blank');
    }
  };

  return (
    <div className="section-container">
      <h3 className="section-title">格式转换</h3>
      
      <div className="section-content">
        {!canProcess && (
          <div className="requirement-notice">
            <p>⚠️ 请先完成源文件和格式要求文件的解析</p>
          </div>
        )}

        {canProcess && (
          <div className="process-area">
            <div className="ready-info">
              <p>✅ 源文件已解析</p>
              <p>✅ 格式要求文件已解析</p>
              <p>可以开始格式转换</p>
            </div>

            <button
              onClick={handleProcess}
              disabled={processing}
              className="action-button process-button"
            >
              {processing ? '转换中...' : '开始格式转换'}
            </button>
          </div>
        )}

        {status && (
          <div className="status-area">
            <div className="status-info">
              <p>状态: {status}</p>
              {taskId && <p>任务ID: {taskId}</p>}
            </div>
          </div>
        )}

        {downloadUrl && (
          <div className="download-area">
            <div className="completion-indicator">
              ✅ 格式转换完成
            </div>
            <button
              onClick={handleDownload}
              className="action-button download-button"
            >
              下载处理结果
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

export default ProcessSection;