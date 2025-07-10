import React, { useState, useEffect } from 'react';
import './UnitInput.css';

interface UnitInputProps {
  value: number;
  unit: string;
  fieldType: string; // 'margin', 'fontSize', 'lineSpacing', etc.
  onChange: (value: number, unit: string) => void;
  availableUnits?: string[];
  showEquivalent?: boolean;
  label?: string;
  min?: number;
  max?: number;
  step?: number;
}

// 单位转换函数
const convertBetweenUnits = (value: number, fromUnit: string, toUnit: string): number => {
  // 先转换为pt
  let ptValue = value;
  
  const toPt: { [key: string]: number } = {
    'pt': 1,
    'px': 0.75,
    'cm': 28.34646,
    'mm': 2.834646,
    'inch': 72,
    'em': 12,
    '字符': 12
  };
  
  if (fromUnit in toPt) {
    ptValue = value * toPt[fromUnit];
  }
  
  // 从pt转换到目标单位
  if (toUnit === 'pt') return ptValue;
  if (toUnit in toPt) return ptValue / toPt[toUnit];
  
  return value;
};

// 获取等价显示
const getEquivalentDisplay = (value: number, unit: string, fieldType: string): string => {
  const ptValue = convertBetweenUnits(value, unit, 'pt');
  
  if (fieldType === 'fontSize') {
    // 中文字号映射
    const chineseSizes: { [key: number]: string } = {
      42: '初号',
      36: '小初',
      26: '一号',
      24: '小一',
      22: '二号',
      18: '小二',
      16: '三号',
      15: '小三',
      14: '四号',
      12: '小四',
      10.5: '五号',
      9: '小五',
      7.5: '六号',
      6.5: '小六',
      5.5: '七号',
      5: '八号'
    };
    
    // 查找最接近的中文字号
    for (const [size, name] of Object.entries(chineseSizes)) {
      if (Math.abs(ptValue - Number(size)) < 0.5) {
        return `(${ptValue}pt, ${name})`;
      }
    }
    return `(${ptValue}pt)`;
  }
  
  if (unit === 'pt') {
    const cmValue = Math.round(ptValue / 28.34646 * 100) / 100;
    return `(${cmValue}cm)`;
  } else if (unit === 'cm') {
    return `(${Math.round(ptValue)}pt)`;
  }
  
  return '';
};

const UnitInput: React.FC<UnitInputProps> = ({
  value,
  unit,
  fieldType,
  onChange,
  availableUnits = ['pt', 'cm', 'mm', 'inch'],
  showEquivalent = true,
  label,
  min,
  max,
  step = 0.1
}) => {
  const [localValue, setLocalValue] = useState(value);
  const [localUnit, setLocalUnit] = useState(unit);
  const [warning, setWarning] = useState<string | null>(null);
  
  useEffect(() => {
    setLocalValue(value);
    setLocalUnit(unit);
  }, [value, unit]);
  
  const handleValueChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newValue = parseFloat(e.target.value) || 0;
    setLocalValue(newValue);
    onChange(newValue, localUnit);
    
    // 验证值
    validateValue(newValue, localUnit);
  };
  
  const handleUnitChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newUnit = e.target.value;
    // 转换值到新单位
    const convertedValue = Math.round(convertBetweenUnits(localValue, localUnit, newUnit) * 100) / 100;
    setLocalValue(convertedValue);
    setLocalUnit(newUnit);
    onChange(convertedValue, newUnit);
    
    validateValue(convertedValue, newUnit);
  };
  
  const validateValue = (val: number, un: string) => {
    const ptValue = convertBetweenUnits(val, un, 'pt');
    
    // 根据字段类型设置警告
    if (fieldType === 'margin') {
      if (ptValue < 10 || ptValue > 150) {
        setWarning('页边距通常在0.35-5.3厘米之间');
      } else {
        setWarning(null);
      }
    } else if (fieldType === 'fontSize') {
      if (ptValue < 8 || ptValue > 72) {
        setWarning('字号通常在8-72pt之间');
      } else {
        setWarning(null);
      }
    } else {
      setWarning(null);
    }
  };
  
  return (
    <div className="unit-input-container">
      {label && <label className="unit-input-label">{label}</label>}
      <div className="unit-input-group">
        <input
          type="number"
          value={localValue}
          onChange={handleValueChange}
          min={min}
          max={max}
          step={step}
          className={`unit-input-field ${warning ? 'has-warning' : ''}`}
        />
        <select
          value={localUnit}
          onChange={handleUnitChange}
          className="unit-input-select"
        >
          {availableUnits.map(u => (
            <option key={u} value={u}>{u}</option>
          ))}
        </select>
        {showEquivalent && (
          <span className="unit-input-equivalent">
            {getEquivalentDisplay(localValue, localUnit, fieldType)}
          </span>
        )}
      </div>
      {warning && <div className="unit-input-warning">{warning}</div>}
    </div>
  );
};

export default UnitInput;