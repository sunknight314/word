import React, { useState } from 'react';
import './StepComponents.css';

interface Step5Props {
  sourceFileId: string | null;
  formatFileId: string | null;
  sourceFileName?: string;
  formatFileName?: string;
}

interface ConvertResult {
  task_id: string;
  status: string;
  message: string;
  output_path?: string;
  report?: {
    total_paragraphs: number;
    styled_paragraphs: number;
    type_distribution: Record<string, number>;
    undefined_styles: string[];
  };
}

const Step5FormatConvert: React.FC<Step5Props> = ({ 
  sourceFileId, 
  formatFileId,
  sourceFileName,
  formatFileName 
}) => {
  const [converting, setConverting] = useState(false);
  const [result, setResult] = useState<ConvertResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleConvert = async () => {
    if (!sourceFileId || !formatFileId) {
      setError('请先完成前面的步骤');
      return;
    }

    setConverting(true);
    setError(null);
    
    try {
      const response = await fetch('/api/process', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          source_id: sourceFileId,
          format_id: formatFileId
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || '格式转换失败');
      }

      const data = await response.json();
      setResult(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : '格式转换失败');
    } finally {
      setConverting(false);
    }
  };

  const canConvert = sourceFileId && formatFileId;

  return (
    <div className="step-container">
      <h3>步骤 5: 执行格式转换</h3>
      <div className="step-content">
        {/* 显示已选择的文件 */}
        <div className="selected-files">
          <div className="file-info-item">
            <span className="file-label">源文档:</span>
            <span className="file-name">{sourceFileName || '未选择'}</span>
          </div>
          <div className="file-info-item">
            <span className="file-label">格式要求:</span>
            <span className="file-name">{formatFileName || '未选择'}</span>
          </div>
        </div>

        {/* 转换按钮 */}
        <button
          onClick={handleConvert}
          disabled={!canConvert || converting}
          className="step-button primary"
        >
          {converting ? '正在转换...' : '开始格式转换'}
        </button>

        {/* 错误信息 */}
        {error && (
          <div className="error-message">
            ❌ {error}
          </div>
        )}

        {/* 转换结果 */}
        {result && (
          <div className="convert-result">
            <div className="result-status success">
              ✅ {result.message}
            </div>

            {/* 处理报告 */}
            {result.report && (
              <div className="process-report">
                <h4>处理报告</h4>
                <div className="report-stats">
                  <div className="stat-item">
                    <span className="stat-label">总段落数:</span>
                    <span className="stat-value">{result.report.total_paragraphs}</span>
                  </div>
                  <div className="stat-item">
                    <span className="stat-label">已格式化段落:</span>
                    <span className="stat-value">{result.report.styled_paragraphs}</span>
                  </div>
                  <div className="stat-item">
                    <span className="stat-label">格式化率:</span>
                    <span className="stat-value">
                      {Math.round((result.report.styled_paragraphs / result.report.total_paragraphs) * 100)}%
                    </span>
                  </div>
                </div>

                {/* 类型分布 */}
                {result.report.type_distribution && (
                  <div className="type-distribution">
                    <h5>段落类型分布</h5>
                    <div className="type-list">
                      {Object.entries(result.report.type_distribution).map(([type, count]) => (
                        <div key={type} className="type-item">
                          <span className="type-name">{type}:</span>
                          <span className="type-count">{count}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* 未定义样式警告 */}
                {result.report.undefined_styles && result.report.undefined_styles.length > 0 && (
                  <div className="undefined-styles-warning">
                    <h5>⚠️ 以下样式未在格式配置中定义:</h5>
                    <ul>
                      {result.report.undefined_styles.map(style => (
                        <li key={style}>{style}</li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            )}

            {/* 下载链接 */}
            {result.output_path && (
              <div className="download-section">
                <h4>下载格式化后的文档</h4>
                <p className="output-path">输出路径: {result.output_path}</p>
                <button className="download-button">
                  📥 下载文档
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default Step5FormatConvert;