import React, { useState } from 'react';
import './StepComponents.css';

interface Step5Props {
  sourceFileId: string;
  formatFileId: string;
}

const Step5Process: React.FC<Step5Props> = ({ sourceFileId, formatFileId }) => {
  const [processing, setProcessing] = useState(false);
  const [taskId, setTaskId] = useState<string | null>(null);
  const [status, setStatus] = useState<string>('');
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const handleProcess = async () => {
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
    <div className="step-container">
      <h3>步骤 5: 格式转换与下载</h3>
      <div className="step-content">
        <button
          onClick={handleProcess}
          disabled={processing}
          className="step-button"
        >
          {processing ? '处理中...' : '开始格式转换'}
        </button>

        {status && (
          <div className="status-info">
            <p>状态: {status}</p>
            {taskId && <p>任务ID: {taskId}</p>}
          </div>
        )}

        {downloadUrl && (
          <button
            onClick={handleDownload}
            className="step-button download-button"
          >
            下载处理结果
          </button>
        )}
      </div>
    </div>
  );
};

export default Step5Process;