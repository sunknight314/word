import React, { useState, useEffect } from 'react';
import PageSettingsEditor from './PageSettingsEditor';
import StyleEditor from './StyleEditor';
import { jsonToFormData, formDataToJson } from '../utils/formatConfigConverter';
import './FormatConfigEditor.css';

interface FormatConfigEditorProps {
  initialConfig: any;
  fileId: string;
  onSave?: (config: any) => void;
  onCancel?: () => void;
}

const FormatConfigEditor: React.FC<FormatConfigEditorProps> = ({
  initialConfig,
  fileId,
  onSave,
  onCancel
}) => {
  const [formData, setFormData] = useState(() => jsonToFormData(initialConfig));
  const [hasChanges, setHasChanges] = useState(false);
  const [saving, setSaving] = useState(false);
  const [errors, setErrors] = useState<string[]>([]);
  const [warnings, setWarnings] = useState<string[]>([]);
  
  // 监听表单变化
  useEffect(() => {
    const currentJson = JSON.stringify(formDataToJson(formData));
    const initialJson = JSON.stringify(initialConfig);
    setHasChanges(currentJson !== initialJson);
  }, [formData, initialConfig]);
  
  const handlePageSettingsChange = (newSettings: any) => {
    setFormData({
      ...formData,
      pageSettings: newSettings
    });
  };
  
  const handleStyleChange = (styleName: string, newConfig: any) => {
    setFormData({
      ...formData,
      styles: {
        ...formData.styles,
        [styleName]: newConfig
      }
    });
  };
  
  const handleSave = async () => {
    setSaving(true);
    setErrors([]);
    setWarnings([]);
    
    try {
      const jsonConfig = formDataToJson(formData);
      
      // 调用API保存
      const response = await fetch(`/api/format-config/${fileId}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(jsonConfig)
      });
      
      const result = await response.json();
      
      if (response.ok) {
        // 保存成功
        if (result.warnings) {
          setWarnings(result.warnings);
        }
        
        // 通知父组件
        if (onSave) {
          onSave(jsonConfig);
        }
        
        // 显示成功消息
        alert('格式配置已保存');
        setHasChanges(false);
      } else {
        // 保存失败
        if (result.detail?.errors) {
          setErrors(result.detail.errors);
        } else {
          setErrors([result.detail || '保存失败']);
        }
      }
    } catch (error: any) {
      setErrors(['保存失败: ' + (error.message || '未知错误')]);
    } finally {
      setSaving(false);
    }
  };
  
  const handleReset = () => {
    if (hasChanges && !window.confirm('确定要重置所有修改吗？')) {
      return;
    }
    setFormData(jsonToFormData(initialConfig));
    setErrors([]);
    setWarnings([]);
  };
  
  const handleCancel = () => {
    if (hasChanges && !window.confirm('有未保存的修改，确定要退出吗？')) {
      return;
    }
    if (onCancel) {
      onCancel();
    }
  };
  
  return (
    <div className="format-config-editor">
      <div className="editor-header">
        <h2 className="editor-title">📝 格式配置编辑器</h2>
        <div className="editor-actions">
          <button
            className="btn btn-primary"
            onClick={handleSave}
            disabled={!hasChanges || saving}
          >
            {saving ? '保存中...' : '保存修改'}
          </button>
          <button
            className="btn btn-secondary"
            onClick={handleReset}
            disabled={!hasChanges}
          >
            重置
          </button>
          <button
            className="btn btn-secondary"
            onClick={handleCancel}
          >
            取消
          </button>
        </div>
      </div>
      
      {errors.length > 0 && (
        <div className="alert alert-error">
          <strong>错误：</strong>
          <ul>
            {errors.map((error, index) => (
              <li key={index}>{error}</li>
            ))}
          </ul>
        </div>
      )}
      
      {warnings.length > 0 && (
        <div className="alert alert-warning">
          <strong>警告：</strong>
          <ul>
            {warnings.map((warning, index) => (
              <li key={index}>{warning}</li>
            ))}
          </ul>
        </div>
      )}
      
      <div className="editor-content">
        <PageSettingsEditor
          settings={formData.pageSettings}
          onChange={handlePageSettingsChange}
        />
        
        <div className="styles-section">
          <h3 className="section-title">样式设置</h3>
          <div className="styles-list">
            {Object.entries(formData.styles).map(([styleName, styleConfig]) => (
              <StyleEditor
                key={styleName}
                styleName={styleName}
                styleConfig={styleConfig as any}
                onChange={handleStyleChange}
              />
            ))}
          </div>
        </div>
      </div>
      
      {hasChanges && (
        <div className="unsaved-changes-indicator">
          <span>有未保存的修改</span>
        </div>
      )}
    </div>
  );
};

export default FormatConfigEditor;