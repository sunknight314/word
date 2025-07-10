import React, { useState, useEffect } from 'react';
import './FontSizeInput.css';

interface FontSizeInputProps {
  value: number; // 总是以pt为单位
  onChange: (value: number) => void;
  label?: string;
}

// 中文字号映射
const CHINESE_FONT_SIZES = [
  { name: '初号', pt: 42 },
  { name: '小初', pt: 36 },
  { name: '一号', pt: 26 },
  { name: '小一', pt: 24 },
  { name: '二号', pt: 22 },
  { name: '小二', pt: 18 },
  { name: '三号', pt: 16 },
  { name: '小三', pt: 15 },
  { name: '四号', pt: 14 },
  { name: '小四', pt: 12 },
  { name: '五号', pt: 10.5 },
  { name: '小五', pt: 9 },
  { name: '六号', pt: 7.5 },
  { name: '小六', pt: 6.5 },
  { name: '七号', pt: 5.5 },
  { name: '八号', pt: 5 }
];

// 常用字号
const COMMON_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 42, 48, 72];

const FontSizeInput: React.FC<FontSizeInputProps> = ({
  value,
  onChange,
  label = '字号'
}) => {
  const [inputMode, setInputMode] = useState<'dropdown' | 'custom'>('dropdown');
  const [customValue, setCustomValue] = useState(value.toString());
  
  // 查找对应的中文字号
  const findChineseName = (pt: number): string => {
    const found = CHINESE_FONT_SIZES.find(size => Math.abs(size.pt - pt) < 0.5);
    return found ? found.name : '';
  };
  
  const [chineseName, setChineseName] = useState(findChineseName(value));
  
  useEffect(() => {
    setCustomValue(value.toString());
    setChineseName(findChineseName(value));
  }, [value]);
  
  const handleDropdownChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newValue = parseFloat(e.target.value);
    onChange(newValue);
  };
  
  const handleCustomChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const val = e.target.value;
    setCustomValue(val);
    const numVal = parseFloat(val);
    if (!isNaN(numVal) && numVal > 0) {
      onChange(numVal);
    }
  };
  
  const handleModeSwitch = () => {
    setInputMode(inputMode === 'dropdown' ? 'custom' : 'dropdown');
  };
  
  // 检查当前值是否在常用列表中
  const isCommonSize = COMMON_SIZES.includes(value);
  
  return (
    <div className="font-size-input-container">
      <div className="font-size-header">
        <label className="font-size-label">{label}</label>
        <button
          type="button"
          onClick={handleModeSwitch}
          className="mode-switch-btn"
          title={inputMode === 'dropdown' ? '切换到自定义输入' : '切换到下拉选择'}
        >
          {inputMode === 'dropdown' ? '⚙' : '📋'}
        </button>
      </div>
      
      <div className="font-size-input-group">
        {inputMode === 'dropdown' ? (
          <>
            <select
              value={value}
              onChange={handleDropdownChange}
              className="font-size-select"
            >
              <optgroup label="常用字号">
                {COMMON_SIZES.map(size => {
                  const chName = findChineseName(size);
                  return (
                    <option key={size} value={size}>
                      {size}pt {chName && `(${chName})`}
                    </option>
                  );
                })}
              </optgroup>
              {!isCommonSize && (
                <optgroup label="当前值">
                  <option value={value}>{value}pt</option>
                </optgroup>
              )}
            </select>
          </>
        ) : (
          <>
            <input
              type="number"
              value={customValue}
              onChange={handleCustomChange}
              min="1"
              max="100"
              step="0.5"
              className="font-size-input"
              placeholder="输入字号"
            />
            <span className="font-size-unit">pt</span>
          </>
        )}
        
        {chineseName && (
          <span className="chinese-name-tag">{chineseName}</span>
        )}
      </div>
      
      <div className="font-size-hints">
        <span className="hint-item">常用: 小四(12pt)</span>
        <span className="hint-item">三号(16pt)</span>
        <span className="hint-item">二号(22pt)</span>
      </div>
    </div>
  );
};

export default FontSizeInput;