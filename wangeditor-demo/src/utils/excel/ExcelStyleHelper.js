/**
 * Excel样式工具类 - 负责处理Excel单元格样式
 */

// 导出样式助手对象
export default {
  /**
   * 应用单元格样式
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   */
  applyCellStyle(htmlCell, excelCell) {
    if (!htmlCell || !excelCell) {
      return;
    }
    
    // 默认字体样式
    excelCell.font = {
      name: '宋体', // 默认字体
      size: 11,
      bold: false,
      italic: false,
      underline: false,
      color: { argb: 'FF000000' }
    };
    
    try {
      // 获取样式信息
      const inlineStyle = htmlCell.style;
      const computedStyle = window.getComputedStyle(htmlCell);
      
      // 处理字体家族
      const fontFamily = inlineStyle.fontFamily || computedStyle.fontFamily;
      if (fontFamily) {
        // 将网页字体映射到Excel支持的字体
        const mappedFont = this.mapWebFontToExcel(fontFamily);
        excelCell.font.name = mappedFont;
      }
      
      // 应用基本文本样式
      this.applyTextStyles(htmlCell, excelCell, inlineStyle, computedStyle);
      
      // 应用颜色样式
      this.applyColorStyles(htmlCell, excelCell, inlineStyle, computedStyle);
      
      // 应用对齐方式
      this.applyAlignmentStyles(excelCell, inlineStyle, computedStyle);
      
      // 应用边框样式
      this.applyBorderStyles(excelCell);
    } catch (error) {
      console.error('应用单元格样式失败:', error);
      // 使用默认样式
      this.applyBorderStyles(excelCell);
    }
  },
  
  /**
   * 应用文本样式
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   * @param {CSSStyleDeclaration} inlineStyle - 内联样式
   * @param {CSSStyleDeclaration} computedStyle - 计算样式
   */
  applyTextStyles(htmlCell, excelCell, inlineStyle, computedStyle) {
    if (!htmlCell || !excelCell) {
      return;
    }
    
    // 检查是否有带样式的子元素
    let fontSize = '';
    
    const styledElements = htmlCell.querySelectorAll('*[style]');
    Array.from(styledElements).forEach(function(elem) {
      const style = elem.style;
      
      // 检查字体大小
      if (style.fontSize) fontSize = style.fontSize;
      
      // 检查文本装饰
      if (style.fontWeight === 'bold' || parseInt(style.fontWeight) >= 700) {
        excelCell.font.bold = true;
      }
      if (style.fontStyle === 'italic') {
        excelCell.font.italic = true;
      }
      if (style.textDecoration && style.textDecoration.indexOf('underline') !== -1) {
        excelCell.font.underline = true;
      }
    });
    
    // 检查标签本身的样式
    if (htmlCell.tagName === 'TH') {
      excelCell.font.bold = true; // 表头默认粗体
    }
    
    // 检查特定元素类型
    const hasStrong = htmlCell.querySelector('strong, b');
    const hasEm = htmlCell.querySelector('em, i');
    const hasU = htmlCell.querySelector('u');
    
    if (hasStrong) excelCell.font.bold = true;
    if (hasEm) excelCell.font.italic = true;
    if (hasU) excelCell.font.underline = true;
    
    // 字体粗细
    if (!excelCell.font.bold) {
      if (inlineStyle.fontWeight === 'bold' || 
          (computedStyle && computedStyle.fontWeight === 'bold') || 
          (computedStyle && parseInt(computedStyle.fontWeight) >= 700)) {
        excelCell.font.bold = true;
      }
    }
    
    // 字体斜体
    if (!excelCell.font.italic) {
      if (inlineStyle.fontStyle === 'italic' || 
          (computedStyle && computedStyle.fontStyle === 'italic')) {
        excelCell.font.italic = true;
      }
    }
    
    // 下划线
    if (!excelCell.font.underline) {
      const hasUnderline = inlineStyle.textDecoration && inlineStyle.textDecoration.indexOf('underline') !== -1;
      const computedUnderline = computedStyle && computedStyle.textDecoration && 
                               computedStyle.textDecoration.indexOf('underline') !== -1;
      if (hasUnderline || computedUnderline) {
        excelCell.font.underline = true;
      }
    }
    
    // 字体大小
    if (fontSize) {
      excelCell.font.size = Math.round(parseInt(fontSize) * 0.75);
    } else if (inlineStyle.fontSize) {
      excelCell.font.size = Math.round(parseInt(inlineStyle.fontSize) * 0.75);
    } else if (computedStyle && computedStyle.fontSize) {
      excelCell.font.size = Math.round(parseInt(computedStyle.fontSize) * 0.75);
    }
  },
  
  /**
   * 应用颜色样式
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   * @param {CSSStyleDeclaration} inlineStyle - 内联样式
   * @param {CSSStyleDeclaration} computedStyle - 计算样式
   */
  applyColorStyles(htmlCell, excelCell, inlineStyle, computedStyle) {
    if (!htmlCell || !excelCell) {
      return;
    }
    
    // 收集颜色信息
    let textColor = '';
    let backgroundColor = '';
    
    const styledElements = htmlCell.querySelectorAll('*[style]');
    Array.from(styledElements).forEach(function(elem) {
      const style = elem.style;
      if (style.color) textColor = style.color;
      if (style.backgroundColor) backgroundColor = style.backgroundColor;
    });
    
    // 字体颜色
    if (textColor) {
      const argb = this.colorToARGB(textColor);
      excelCell.font.color = { argb: argb };
    } else if (inlineStyle.color) {
      const argb = this.colorToARGB(inlineStyle.color);
      excelCell.font.color = { argb: argb };
    } else if (computedStyle && computedStyle.color && computedStyle.color !== 'rgb(0, 0, 0)') {
      const argb = this.colorToARGB(computedStyle.color);
      excelCell.font.color = { argb: argb };
    }
    
    // 背景色
    if (backgroundColor) {
      const argb = this.colorToARGB(backgroundColor);
      excelCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: argb }
      };
    } else if (inlineStyle.backgroundColor && inlineStyle.backgroundColor !== 'transparent') {
      const argb = this.colorToARGB(inlineStyle.backgroundColor);
      excelCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: argb }
      };
    } else if (computedStyle && computedStyle.backgroundColor && 
              computedStyle.backgroundColor !== 'rgba(0, 0, 0, 0)' && 
              computedStyle.backgroundColor !== 'transparent') {
      const argb = this.colorToARGB(computedStyle.backgroundColor);
      excelCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: argb }
      };
    }
  },
  
  /**
   * 应用对齐方式样式
   * @param {Object} excelCell - Excel单元格对象
   * @param {CSSStyleDeclaration} inlineStyle - 内联样式
   * @param {CSSStyleDeclaration} computedStyle - 计算样式
   */
  applyAlignmentStyles(excelCell, inlineStyle, computedStyle) {
    if (!excelCell) {
      return;
    }
    
    // 对齐方式
    const alignment = {
      wrapText: true, // 自动换行
      vertical: 'middle' // 默认垂直居中
    };
    
    // 水平对齐
    const textAlign = inlineStyle.textAlign || (computedStyle ? computedStyle.textAlign : '');
    if (textAlign) {
      if (textAlign === 'center') {
        alignment.horizontal = 'center';
      } else if (textAlign === 'right') {
        alignment.horizontal = 'right';
      } else if (textAlign === 'left' || textAlign === 'start') {
        alignment.horizontal = 'left';
      }
    }
    
    // 垂直对齐
    const verticalAlign = inlineStyle.verticalAlign || (computedStyle ? computedStyle.verticalAlign : '');
    if (verticalAlign) {
      if (verticalAlign === 'middle') {
        alignment.vertical = 'middle';
      } else if (verticalAlign === 'top') {
        alignment.vertical = 'top';
      } else if (verticalAlign === 'bottom') {
        alignment.vertical = 'bottom';
      }
    }
    
    excelCell.alignment = alignment;
  },
  
  /**
   * 应用边框样式
   * @param {Object} excelCell - Excel单元格对象
   */
  applyBorderStyles(excelCell) {
    if (!excelCell) {
      return;
    }
    
    // 完整的边框样式
    excelCell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };
  },
  
  /**
   * 将颜色值转换为ARGB格式
   * @param {String} color - 颜色值（支持rgb, rgba, hex）
   * @returns {String} - ARGB格式颜色值
   */
  colorToARGB(color) {
    if (!color) {
      return 'FF000000';
    }
    
    // 处理hex格式
    if (color.startsWith('#')) {
      let hex = color.slice(1);
      if (hex.length === 3) {
        hex = hex.split('').map(function(c) { return c + c; }).join('');
      }
      return 'FF' + hex.toUpperCase();
    }
    
    // 处理rgb格式
    const rgbMatch = color.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
    if (rgbMatch) {
      const r = rgbMatch[1];
      const g = rgbMatch[2];
      const b = rgbMatch[3];
      return 'FF' + this.numberToHex(r) + this.numberToHex(g) + this.numberToHex(b);
    }
    
    // 处理rgba格式
    const rgbaMatch = color.match(/^rgba\((\d+),\s*(\d+),\s*(\d+),\s*([0-9.]+)\)$/);
    if (rgbaMatch) {
      const r = rgbaMatch[1];
      const g = rgbaMatch[2];
      const b = rgbaMatch[3];
      const a = rgbaMatch[4];
      const alpha = Math.round(parseFloat(a) * 255).toString(16).toUpperCase().padStart(2, '0');
      return alpha + this.numberToHex(r) + this.numberToHex(g) + this.numberToHex(b);
    }
    
    // 处理颜色名称映射
    return this.getNamedColorARGB(color) || 'FF000000'; // 默认黑色
  },
  
  /**
   * 获取命名颜色的ARGB值
   * @param {String} color - 颜色名称
   * @returns {String} - ARGB格式颜色值
   */
  getNamedColorARGB(color) {
    const colorMap = {
      'black': 'FF000000',
      'white': 'FFFFFFFF',
      'red': 'FFFF0000',
      'green': 'FF008000',
      'blue': 'FF0000FF',
      'yellow': 'FFFFFF00',
      'magenta': 'FFFF00FF',
      'cyan': 'FF00FFFF',
      'gray': 'FF808080',
      'silver': 'FFC0C0C0',
      'purple': 'FF800080',
      'orange': 'FFFFA500',
      'pink': 'FFFFC0CB'
    };
    
    const colorName = color.toLowerCase().trim();
    return colorMap[colorName];
  },
  
  /**
   * 数字转为两位十六进制
   * @param {Number|String} num - 数字
   * @returns {String} - 两位十六进制字符串
   */
  numberToHex(num) {
    const hex = parseInt(num).toString(16).toUpperCase();
    return hex.length === 1 ? '0' + hex : hex;
  },
  
  /**
   * 映射网页字体到Excel支持的字体
   * @param {String} webFont - 网页字体名称
   * @returns {String} - Excel支持的字体名称
   */
  mapWebFontToExcel(webFont) {
    // 如果没有字体，返回默认字体
    if (!webFont) {
      return "宋体";
    }
    
    const fontMap = {
      // 中文字体
      "宋体": "宋体",
      "simsun": "宋体",
      "黑体": "黑体",
      "simhei": "黑体",
      "微软雅黑": "微软雅黑",
      "microsoft yahei": "微软雅黑",
      "楷体": "楷体",
      "simkai": "楷体",
      "仿宋": "仿宋",
      "fangsong": "仿宋",
      
      // 英文字体
      "arial": "Arial",
      "helvetica": "Arial",
      "times new roman": "Times New Roman",
      "calibri": "Calibri",
      "verdana": "Verdana",
      "tahoma": "Tahoma",
      "courier new": "Courier New",
      "georgia": "Georgia",
      
      // 通用字体
      "sans-serif": "Arial",
      "serif": "Times New Roman",
      "monospace": "Courier New"
    };
    
    // 将字体名称转为小写，去除引号和空格
    const normalizedFont = webFont.toLowerCase()
      .replace(/["']/g, '')
      .trim();
    
    // 如果找到映射，返回对应的Excel字体
    return fontMap[normalizedFont] || "宋体"; // 默认返回宋体
  },
  
  /**
   * 计算内容的显示宽度
   * @param {String} content - 内容
   * @returns {Number} - 估算的像素宽度
   */
  calculateContentWidth(content) {
    if (!content) return 0;
    
    // 基础字符宽度（像素）
    const charWidths = {
      default: 7,  // 默认ASCII字符宽度
      wide: 12,    // 宽字符（如中文）宽度
      narrow: 4    // 窄字符宽度
    };
    
    let totalWidth = 0;
    
    // 统计不同类型字符
    for (let i = 0; i < content.length; i++) {
      const char = content.charAt(i);
      const code = char.charCodeAt(0);
      
      // 中文字符和其他全角字符
      if ((code >= 0x4E00 && code <= 0x9FFF) || 
          (code >= 0x3000 && code <= 0x303F) || 
          (code >= 0xFF00 && code <= 0xFFEF)) {
        totalWidth += charWidths.wide;
      } 
      // 窄字符
      else if (char === 'i' || char === 'l' || char === 'I' || char === '.' || char === ',' || char === ':' || char === ';') {
        totalWidth += charWidths.narrow;
      } 
      // 默认ASCII字符
      else {
        totalWidth += charWidths.default;
      }
    }
    
    // 处理换行情况
    const lines = content.split('\n');
    if (lines.length > 1) {
      // 取最长行的宽度
      let maxLineWidth = 0;
      for (let i = 0; i < lines.length; i++) {
        const lineWidth = this.calculateContentWidth(lines[i]);
        if (lineWidth > maxLineWidth) {
          maxLineWidth = lineWidth;
        }
      }
      totalWidth = maxLineWidth;
    }
    
    return totalWidth;
  }
};
