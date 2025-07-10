// 单位转换工具类

export class UnitConverter {
  // 中文字号映射表
  private chineseFontSizeMap: { [key: string]: number } = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5
  };

  // 反向映射：pt值到中文字号
  private ptToChineseSizeMap: { [key: number]: string } = {};

  constructor() {
    // 创建反向映射
    Object.entries(this.chineseFontSizeMap).forEach(([name, pt]) => {
      this.ptToChineseSizeMap[pt] = name;
    });
  }

  // 转换为pt
  convertToPt(value: number, unit: string): number {
    switch (unit.toLowerCase()) {
      case 'pt':
        return value;
      case 'px':
        return value * 0.75;
      case 'cm':
        return value * 28.35;
      case 'mm':
        return value * 2.835;
      case 'inch':
      case 'in':
        return value * 72;
      default:
        // 检查是否是中文字号
        if (this.chineseFontSizeMap[unit]) {
          return this.chineseFontSizeMap[unit];
        }
        return value;
    }
  }

  // 从pt转换到其他单位
  convertFromPt(ptValue: number, targetUnit: string): number {
    switch (targetUnit.toLowerCase()) {
      case 'pt':
        return ptValue;
      case 'px':
        return ptValue / 0.75;
      case 'cm':
        return ptValue / 28.35;
      case 'mm':
        return ptValue / 2.835;
      case 'inch':
      case 'in':
        return ptValue / 72;
      default:
        return ptValue;
    }
  }

  // 获取中文字号名称
  getChineseFontSizeName(pt: number): string | null {
    // 直接匹配
    if (this.ptToChineseSizeMap[pt]) {
      return this.ptToChineseSizeMap[pt];
    }
    
    // 允许小范围误差（0.5pt）
    for (const [ptSize, name] of Object.entries(this.ptToChineseSizeMap)) {
      if (Math.abs(Number(ptSize) - pt) < 0.5) {
        return name;
      }
    }
    
    return null;
  }

  // 获取字号显示文本
  getFontSizeDisplay(pt: number): string {
    const chineseName = this.getChineseFontSizeName(pt);
    if (chineseName) {
      return `${pt}pt (${chineseName})`;
    }
    return `${pt}pt`;
  }

  // 验证值是否在合理范围内
  validateValue(value: number, unit: string, fieldType: string): { valid: boolean; warning?: string; error?: string } {
    const ptValue = this.convertToPt(value, unit);

    switch (fieldType) {
      case 'margin':
        if (ptValue < 0) {
          return { valid: false, error: '页边距不能为负数' };
        }
        if (ptValue > 300) {
          return { valid: true, warning: '页边距过大，可能导致内容区域过小' };
        }
        break;
      
      case 'fontSize':
        if (ptValue < 5) {
          return { valid: false, error: '字号过小，建议不小于5pt' };
        }
        if (ptValue > 100) {
          return { valid: true, warning: '字号过大，建议不超过72pt' };
        }
        break;
      
      case 'lineSpacing':
        if (ptValue < 0) {
          return { valid: false, error: '行距不能为负数' };
        }
        if (ptValue > 100) {
          return { valid: true, warning: '行距过大，建议不超过3倍字号' };
        }
        break;
      
      case 'spacing':
        if (ptValue < 0) {
          return { valid: false, error: '间距不能为负数' };
        }
        if (ptValue > 200) {
          return { valid: true, warning: '间距过大' };
        }
        break;
    }

    return { valid: true };
  }

  // 格式化数值显示
  formatValue(value: number, decimals: number = 2): string {
    return value.toFixed(decimals).replace(/\.?0+$/, '');
  }
}

export const unitConverter = new UnitConverter();