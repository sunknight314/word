/**
 * 格式配置转换工具
 * 负责JSON配置和表单数据之间的双向转换
 */

export interface DisplayValue {
  value: number;
  unit: string;
  ptValue: number;
  display: string;
}

export interface FormMargins {
  top: DisplayValue;
  bottom: DisplayValue;
  left: DisplayValue;
  right: DisplayValue;
}

export interface FormPageSettings {
  paperSize: string;
  orientation: string;
  margins: FormMargins;
}

export interface FormStyleConfig {
  fontSizePt: number;
  fontNameCn: string;
  fontNameEn: string;
  alignment: string;
  lineSpacing: DisplayValue;
  spaceBefore?: DisplayValue;
  spaceAfter?: DisplayValue;
  firstLineIndent?: DisplayValue;
  bold?: boolean;
}

export interface FormData {
  pageSettings: FormPageSettings;
  styles: {
    [key: string]: FormStyleConfig;
  };
}

// 从pt字符串提取数值
const extractPtValue = (value: string): number => {
  if (typeof value !== 'string') return 0;
  const match = value.match(/^([\d.]+)\s*pt$/);
  return match ? parseFloat(match[1]) : 0;
};

// 将pt值转换为显示值
const ptToDisplayValue = (ptValue: number, fieldType: string): DisplayValue => {
  let unit = 'pt';
  let displayValue = ptValue;
  
  if (fieldType === 'margin' || fieldType === 'spacing' || fieldType === 'indent') {
    // 边距类默认显示为cm
    unit = 'cm';
    displayValue = Math.round(ptValue / 28.34646 * 100) / 100;
  } else if (fieldType === 'lineSpacing' && ptValue % 12 === 0 && ptValue > 0) {
    // 行距如果是12的倍数，显示为倍数
    unit = '倍';
    displayValue = ptValue / 12;
  }
  
  return {
    value: displayValue,
    unit,
    ptValue,
    display: `${displayValue}${unit}`
  };
};

// 将显示值转换为pt字符串
const displayValueToPt = (displayValue: DisplayValue): string => {
  return `${displayValue.ptValue}pt`;
};

/**
 * 将JSON格式配置转换为表单数据
 */
export const jsonToFormData = (jsonConfig: any): FormData => {
  const formData: FormData = {
    pageSettings: {
      paperSize: jsonConfig.page_settings?.paper_size || 'A4',
      orientation: jsonConfig.page_settings?.orientation || 'portrait',
      margins: {
        top: ptToDisplayValue(extractPtValue(jsonConfig.page_settings?.margins?.top || '72pt'), 'margin'),
        bottom: ptToDisplayValue(extractPtValue(jsonConfig.page_settings?.margins?.bottom || '72pt'), 'margin'),
        left: ptToDisplayValue(extractPtValue(jsonConfig.page_settings?.margins?.left || '72pt'), 'margin'),
        right: ptToDisplayValue(extractPtValue(jsonConfig.page_settings?.margins?.right || '72pt'), 'margin')
      }
    },
    styles: {}
  };
  
  // 转换样式
  if (jsonConfig.styles) {
    for (const [styleName, styleConfig] of Object.entries(jsonConfig.styles)) {
      const config = styleConfig as any;
      const style: FormStyleConfig = {
        fontSizePt: extractPtValue(config.font_size || '12pt'),
        fontNameCn: config.font_name_cn || '宋体',
        fontNameEn: config.font_name_en || 'Times New Roman',
        alignment: config.alignment || 'left',
        lineSpacing: ptToDisplayValue(extractPtValue(config.line_spacing || '20pt'), 'lineSpacing'),
        bold: config.bold === true || config.bold === 'true'
      };
      
      // 可选字段
      if (config.space_before) {
        style.spaceBefore = ptToDisplayValue(extractPtValue(config.space_before), 'spacing');
      }
      if (config.space_after) {
        style.spaceAfter = ptToDisplayValue(extractPtValue(config.space_after), 'spacing');
      }
      if (config.first_line_indent) {
        style.firstLineIndent = ptToDisplayValue(extractPtValue(config.first_line_indent), 'indent');
      }
      
      formData.styles[styleName] = style;
    }
  }
  
  return formData;
};

/**
 * 将表单数据转换为JSON格式配置
 */
export const formDataToJson = (formData: FormData): any => {
  const jsonConfig: any = {
    page_settings: {
      paper_size: formData.pageSettings.paperSize,
      orientation: formData.pageSettings.orientation,
      margins: {
        top: displayValueToPt(formData.pageSettings.margins.top),
        bottom: displayValueToPt(formData.pageSettings.margins.bottom),
        left: displayValueToPt(formData.pageSettings.margins.left),
        right: displayValueToPt(formData.pageSettings.margins.right)
      }
    },
    styles: {}
  };
  
  // 转换样式
  for (const [styleName, styleConfig] of Object.entries(formData.styles)) {
    const style: any = {
      font_size: `${styleConfig.fontSizePt}pt`,
      font_name_cn: styleConfig.fontNameCn,
      font_name_en: styleConfig.fontNameEn,
      alignment: styleConfig.alignment,
      line_spacing: displayValueToPt(styleConfig.lineSpacing)
    };
    
    // 可选字段
    if (styleConfig.bold) {
      style.bold = true;
    }
    if (styleConfig.spaceBefore) {
      style.space_before = displayValueToPt(styleConfig.spaceBefore);
    }
    if (styleConfig.spaceAfter) {
      style.space_after = displayValueToPt(styleConfig.spaceAfter);
    }
    if (styleConfig.firstLineIndent) {
      style.first_line_indent = displayValueToPt(styleConfig.firstLineIndent);
    }
    
    jsonConfig.styles[styleName] = style;
  }
  
  return jsonConfig;
};

/**
 * 验证规则
 */
export const validationRules = {
  margins: {
    min: 0,
    max: 300,
    warningMin: 10,
    warningMax: 150,
    message: '页边距通常在0.35-5.3厘米之间'
  },
  fontSize: {
    min: 5,
    max: 100,
    warningMin: 8,
    warningMax: 72,
    message: '字号通常在8-72pt之间'
  },
  lineSpacing: {
    min: 0.5,
    max: 10,
    warningMin: 1,
    warningMax: 3,
    message: '行距通常在1-3倍之间'
  },
  spacing: {
    min: 0,
    max: 200,
    warningMin: 0,
    warningMax: 50,
    message: '段间距通常在0-50pt之间'
  },
  indent: {
    min: 0,
    max: 200,
    warningMin: 0,
    warningMax: 72,
    message: '缩进通常在0-72pt之间'
  }
};

/**
 * 单位转换辅助函数
 */
export const convertUnit = (value: number, fromUnit: string, toUnit: string): number => {
  // 转换系数（到pt）
  const toPt: { [key: string]: number } = {
    'pt': 1,
    'px': 0.75,
    'cm': 28.34646,
    'mm': 2.834646,
    'inch': 72,
    'em': 12,
    '字符': 12,
    '倍': 12 // 行距倍数
  };
  
  // 先转为pt
  const ptValue = value * (toPt[fromUnit] || 1);
  
  // 再转为目标单位
  return ptValue / (toPt[toUnit] || 1);
};