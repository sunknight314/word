import React, { useState } from 'react';
import './SectionComponents.css';

interface ProcessSectionProps {
  sourceFileId: string;
  formatFileId: string;
  canProcess: boolean;
}

interface ProcessResult {
  task_id: string;
  status: string;
  message: string;
  output_path?: string;
  steps_results?: {
    page_settings?: any;
    create_styles?: any;
    apply_styles?: any;
    sections?: any;
    toc?: any;
    save?: any;
  };
  report?: {
    total_steps?: number;
    completed_steps?: number;
    summary?: {
      styles_created?: number;
      styles_updated?: number;
      paragraphs_styled?: number;
      toc_created?: boolean;
      page_numbers_added?: boolean;
      headers_footers_added?: boolean;
    };
    total_paragraphs?: number;
    styled_paragraphs?: number;
    type_distribution?: Record<string, number>;
    undefined_styles?: string[];
  };
}

const ProcessSection: React.FC<ProcessSectionProps> = ({ 
  sourceFileId, 
  formatFileId, 
  canProcess 
}) => {
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState<ProcessResult | null>(null);
  const [error, setError] = useState<string | null>(null);

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

      if (!processResponse.ok) {
        const errorData = await processResponse.json();
        throw new Error(errorData.detail || '格式转换失败');
      }

      const processResult: ProcessResult = await processResponse.json();
      setResult(processResult);
      setProcessing(false);
    } catch (err) {
      console.error('处理失败:', err);
      setError(err instanceof Error ? err.message : '格式转换失败');
      setProcessing(false);
    }
  };

  const handleDownload = () => {
    if (result?.output_path) {
      // 创建下载链接
      const downloadUrl = `/api/download/${encodeURIComponent(result.output_path)}`;
      window.open(downloadUrl, '_blank');
    }
  };

  return (
    <div className="section-container">
      <h3 className="section-title">格式转换</h3>
      
      <div className="section-content">
        <div className="process-status">
          {!canProcess ? (
            <div className="requirement-notice">
              <p>⚠️ 请先完成源文件和格式要求文件的解析</p>
            </div>
          ) : (
            <div className="ready-info">
              <p>✅ 源文件已解析</p>
              <p>✅ 格式要求文件已解析</p>
              <p>可以开始格式转换</p>
            </div>
          )}
        </div>

        <div className="process-button-area">
          <button
            onClick={handleProcess}
            disabled={!canProcess || processing}
            className={`action-button process-button ${canProcess ? 'ready' : 'disabled'}`}
          >
            {processing ? '转换中...' : '开始格式转换'}
          </button>
        </div>

        {error && (
          <div className="error-message">
            ❌ {error}
          </div>
        )}

        {result && (
          <div className="result-area">
            <div className="completion-indicator">
              ✅ {result.message}
            </div>
            
            {/* 处理报告 */}
            {result.report && (
              <div className="process-report">
                <h4>处理报告</h4>
                
                {/* 步骤完成情况 */}
                {result.report.total_steps && (
                  <div className="steps-summary">
                    <p>处理步骤: {result.report.completed_steps} / {result.report.total_steps} 完成</p>
                  </div>
                )}
                
                {/* 详细统计 */}
                {result.report.summary && (
                  <div className="report-stats">
                    <h5>处理详情</h5>
                    {result.report.summary.styles_created !== undefined && (
                      <p>✅ 创建样式: {result.report.summary.styles_created} 个</p>
                    )}
                    {result.report.summary.styles_updated !== undefined && (
                      <p>✅ 更新样式: {result.report.summary.styles_updated} 个</p>
                    )}
                    {result.report.summary.paragraphs_styled !== undefined && (
                      <p>✅ 应用样式段落: {result.report.summary.paragraphs_styled} 个</p>
                    )}
                    {result.report.summary.toc_created && (
                      <p>✅ 已创建目录</p>
                    )}
                    {result.report.summary.page_numbers_added && (
                      <p>✅ 已添加页码</p>
                    )}
                    {result.report.summary.headers_footers_added && (
                      <p>✅ 已添加页眉页脚</p>
                    )}
                  </div>
                )}
                
                {/* 旧版报告兼容 */}
                {result.report.total_paragraphs && (
                  <div className="report-stats">
                    <p>总段落数: {result.report.total_paragraphs}</p>
                    <p>已格式化: {result.report.styled_paragraphs}</p>
                    <p>格式化率: {Math.round((result.report.styled_paragraphs! / result.report.total_paragraphs) * 100)}%</p>
                  </div>
                )}
                
                {result.report.undefined_styles && result.report.undefined_styles.length > 0 && (
                  <div className="warning-info">
                    <p>⚠️ 以下样式未定义:</p>
                    <ul>
                      {result.report.undefined_styles.map(style => (
                        <li key={style}>{style}</li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            )}
            
            {result.output_path && (
              <div className="download-area">
                <button
                  onClick={handleDownload}
                  className="action-button download-button"
                >
                  下载处理结果
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default ProcessSection;