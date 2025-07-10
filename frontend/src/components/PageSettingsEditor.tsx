import React from 'react';
import UnitInput from './UnitInput';
import { FormPageSettings } from '../utils/formatConfigConverter';
import './PageSettingsEditor.css';

interface PageSettingsEditorProps {
  settings: FormPageSettings;
  onChange: (settings: FormPageSettings) => void;
}

const PageSettingsEditor: React.FC<PageSettingsEditorProps> = ({
  settings,
  onChange
}) => {
  const handleMarginChange = (
    side: 'top' | 'bottom' | 'left' | 'right',
    value: number,
    unit: string
  ) => {
    const ptValue = convertToPt(value, unit);
    onChange({
      ...settings,
      margins: {
        ...settings.margins,
        [side]: {
          value,
          unit,
          ptValue,
          display: `${value}${unit}`
        }
      }
    });
  };
  
  const handlePaperSizeChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    onChange({
      ...settings,
      paperSize: e.target.value
    });
  };
  
  const handleOrientationChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    onChange({
      ...settings,
      orientation: e.target.value
    });
  };
  
  // 单位转换辅助函数
  const convertToPt = (value: number, unit: string): number => {
    const conversions: { [key: string]: number } = {
      'pt': 1,
      'cm': 28.34646,
      'mm': 2.834646,
      'inch': 72
    };
    return value * (conversions[unit] || 1);
  };
  
  return (
    <div className="page-settings-editor">
      <h3 className="section-title">页面设置</h3>
      
      <div className="settings-group">
        <div className="setting-row">
          <label className="setting-label">纸张大小</label>
          <select
            value={settings.paperSize}
            onChange={handlePaperSizeChange}
            className="setting-select"
          >
            <option value="A3">A3</option>
            <option value="A4">A4</option>
            <option value="A5">A5</option>
            <option value="B4">B4</option>
            <option value="B5">B5</option>
            <option value="Letter">Letter</option>
            <option value="Legal">Legal</option>
          </select>
        </div>
        
        <div className="setting-row">
          <label className="setting-label">页面方向</label>
          <div className="radio-group">
            <label className="radio-label">
              <input
                type="radio"
                name="orientation"
                value="portrait"
                checked={settings.orientation === 'portrait'}
                onChange={handleOrientationChange}
              />
              <span>纵向</span>
            </label>
            <label className="radio-label">
              <input
                type="radio"
                name="orientation"
                value="landscape"
                checked={settings.orientation === 'landscape'}
                onChange={handleOrientationChange}
              />
              <span>横向</span>
            </label>
          </div>
        </div>
      </div>
      
      <div className="margins-section">
        <h4 className="subsection-title">页边距</h4>
        <div className="margins-grid">
          <UnitInput
            label="上边距"
            value={settings.margins.top.value}
            unit={settings.margins.top.unit}
            fieldType="margin"
            onChange={(value, unit) => handleMarginChange('top', value, unit)}
            availableUnits={['cm', 'mm', 'inch', 'pt']}
          />
          <UnitInput
            label="下边距"
            value={settings.margins.bottom.value}
            unit={settings.margins.bottom.unit}
            fieldType="margin"
            onChange={(value, unit) => handleMarginChange('bottom', value, unit)}
            availableUnits={['cm', 'mm', 'inch', 'pt']}
          />
          <UnitInput
            label="左边距"
            value={settings.margins.left.value}
            unit={settings.margins.left.unit}
            fieldType="margin"
            onChange={(value, unit) => handleMarginChange('left', value, unit)}
            availableUnits={['cm', 'mm', 'inch', 'pt']}
          />
          <UnitInput
            label="右边距"
            value={settings.margins.right.value}
            unit={settings.margins.right.unit}
            fieldType="margin"
            onChange={(value, unit) => handleMarginChange('right', value, unit)}
            availableUnits={['cm', 'mm', 'inch', 'pt']}
          />
        </div>
      </div>
    </div>
  );
};

export default PageSettingsEditor;