import React, { useState } from 'react';
import UnitInput from './UnitInput';
import FontSizeInput from './FontSizeInput';
import { FormStyleConfig, convertUnit } from '../utils/formatConfigConverter';
import './StyleEditor.css';

interface StyleEditorProps {
  styleName: string;
  styleConfig: FormStyleConfig;
  onChange: (styleName: string, config: FormStyleConfig) => void;
}

// 样式名称映射
const STYLE_NAME_MAP: { [key: string]: string } = {
  'title': '文档标题',
  'heading1': '一级标题',
  'heading2': '二级标题',
  'heading3': '三级标题',
  'heading4': '四级标题',
  'paragraph': '正文段落',
  'abstract_cn': '中文摘要',
  'abstract_en': '英文摘要',
  'keywords_cn': '中文关键词',
  'keywords_en': '英文关键词',
  'figure_caption': '图注',
  'table_caption': '表注',
  'reference': '参考文献'
};

const StyleEditor: React.FC<StyleEditorProps> = ({
  styleName,
  styleConfig,
  onChange
}) => {
  const [isExpanded, setIsExpanded] = useState(false);
  
  const displayName = STYLE_NAME_MAP[styleName] || styleName;
  
  const handleChange = (field: keyof FormStyleConfig, value: any) => {
    onChange(styleName, {
      ...styleConfig,
      [field]: value
    });
  };
  
  const handleUnitInputChange = (
    field: 'lineSpacing' | 'spaceBefore' | 'spaceAfter' | 'firstLineIndent',
    value: number,
    unit: string
  ) => {
    const ptValue = convertToPt(value, unit);
    handleChange(field, {
      value,
      unit,
      ptValue,
      display: `${value}${unit}`
    });
  };
  
  const convertToPt = (value: number, unit: string): number => {
    return convertUnit(value, unit, 'pt');
  };
  
  return (
    <div className="style-editor">
      <div className="style-header" onClick={() => setIsExpanded(!isExpanded)}>
        <span className="expand-icon">{isExpanded ? '▼' : '▶'}</span>
        <h4 className="style-title">{displayName}</h4>
        <div className="style-summary">
          <span className="summary-item">{styleConfig.fontNameCn}</span>
          <span className="summary-item">{styleConfig.fontSizePt}pt</span>
          <span className="summary-item">{styleConfig.alignment === 'center' ? '居中' : styleConfig.alignment === 'justify' ? '两端对齐' : '左对齐'}</span>
        </div>
      </div>
      
      {isExpanded && (
        <div className="style-content">
          <div className="style-section">
            <h5 className="style-section-title">字体设置</h5>
            
            <div className="form-row">
              <label className="form-label">中文字体</label>
              <select
                value={styleConfig.fontNameCn}
                onChange={(e) => handleChange('fontNameCn', e.target.value)}
                className="form-select"
              >
                <option value="宋体">宋体</option>
                <option value="黑体">黑体</option>
                <option value="仿宋">仿宋</option>
                <option value="楷体">楷体</option>
                <option value="微软雅黑">微软雅黑</option>
                <option value="华文中宋">华文中宋</option>
                <option value="华文楷体">华文楷体</option>
              </select>
            </div>
            
            <div className="form-row">
              <label className="form-label">英文字体</label>
              <select
                value={styleConfig.fontNameEn}
                onChange={(e) => handleChange('fontNameEn', e.target.value)}
                className="form-select"
              >
                <option value="Times New Roman">Times New Roman</option>
                <option value="Arial">Arial</option>
                <option value="Calibri">Calibri</option>
                <option value="Helvetica">Helvetica</option>
                <option value="Georgia">Georgia</option>
                <option value="Cambria">Cambria</option>
              </select>
            </div>
            
            <FontSizeInput
              value={styleConfig.fontSizePt}
              onChange={(value) => handleChange('fontSizePt', value)}
              label="字号"
            />
            
            <div className="form-row">
              <label className="form-label">
                <input
                  type="checkbox"
                  checked={styleConfig.bold || false}
                  onChange={(e) => handleChange('bold', e.target.checked)}
                />
                <span>加粗</span>
              </label>
            </div>
          </div>
          
          <div className="style-section">
            <h5 className="style-section-title">段落设置</h5>
            
            <div className="form-row">
              <label className="form-label">对齐方式</label>
              <select
                value={styleConfig.alignment}
                onChange={(e) => handleChange('alignment', e.target.value)}
                className="form-select"
              >
                <option value="left">左对齐</option>
                <option value="center">居中</option>
                <option value="right">右对齐</option>
                <option value="justify">两端对齐</option>
              </select>
            </div>
            
            <UnitInput
              label="行距"
              value={styleConfig.lineSpacing.value}
              unit={styleConfig.lineSpacing.unit}
              fieldType="lineSpacing"
              onChange={(value, unit) => handleUnitInputChange('lineSpacing', value, unit)}
              availableUnits={['pt', '倍']}
            />
            
            {styleConfig.spaceBefore !== undefined && (
              <UnitInput
                label="段前距"
                value={styleConfig.spaceBefore.value}
                unit={styleConfig.spaceBefore.unit}
                fieldType="spacing"
                onChange={(value, unit) => handleUnitInputChange('spaceBefore', value, unit)}
                availableUnits={['pt', 'cm', 'mm']}
              />
            )}
            
            {styleConfig.spaceAfter !== undefined && (
              <UnitInput
                label="段后距"
                value={styleConfig.spaceAfter.value}
                unit={styleConfig.spaceAfter.unit}
                fieldType="spacing"
                onChange={(value, unit) => handleUnitInputChange('spaceAfter', value, unit)}
                availableUnits={['pt', 'cm', 'mm']}
              />
            )}
            
            {styleConfig.firstLineIndent !== undefined && (
              <UnitInput
                label="首行缩进"
                value={styleConfig.firstLineIndent.value}
                unit={styleConfig.firstLineIndent.unit}
                fieldType="indent"
                onChange={(value, unit) => handleUnitInputChange('firstLineIndent', value, unit)}
                availableUnits={['pt', 'cm', '字符']}
              />
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default StyleEditor;