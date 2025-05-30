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
      this.applyTextStylesFromElements(htmlCell, excelCell);
      this.applyTextStylesFromCSS(htmlCell, excelCell, inlineStyle, computedStyle);
      this.applyFontSize(htmlCell, excelCell, inlineStyle, computedStyle);
      
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
   * 从HTML元素标签应用文本样式
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   */
  applyTextStylesFromElements(htmlCell, excelCell) {
    if (!htmlCell || !excelCell) {
      return;
    }
    
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
  },
  
  /**
   * 从内联子元素收集样式信息
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   */
  applyTextStylesFromInlineElements(htmlCell, excelCell) {
    if (!htmlCell || !excelCell) {
      return;
    }
    
    // 检查是否有带样式的子元素
    const styledElements = htmlCell.querySelectorAll('*[style]');
    
    Array.from(styledElements).forEach(function(elem) {
      const style = elem.style;
      
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
  },
  
  /**
   * 从CSS样式应用文本样式
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   * @param {CSSStyleDeclaration} inlineStyle - 内联样式
   * @param {CSSStyleDeclaration} computedStyle - 计算样式
   */
  applyTextStylesFromCSS(htmlCell, excelCell, inlineStyle, computedStyle) {
    if (!htmlCell || !excelCell || !inlineStyle) {
      return;
    }
    
    // 先应用内联子元素样式
    this.applyTextStylesFromInlineElements(htmlCell, excelCell);
    
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
  },
  
  /**
   * 应用字体大小
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   * @param {CSSStyleDeclaration} inlineStyle - 内联样式
   * @param {CSSStyleDeclaration} computedStyle - 计算样式
   */
  applyFontSize(htmlCell, excelCell, inlineStyle, computedStyle) {
    if (!htmlCell || !excelCell) {
      return;
    }
    
    // 查找所有带有fontSize样式的元素
    let fontSize = this.findFontSizeInChildren(htmlCell);
    
    // 如果子元素中没有找到，检查自身样式
    if (!fontSize && inlineStyle.fontSize) {
      fontSize = inlineStyle.fontSize;
    } 
    // 最后使用计算样式
    else if (!fontSize && computedStyle && computedStyle.fontSize) {
      fontSize = computedStyle.fontSize;
    }
    
    // 应用字体大小
    if (fontSize) {
      // 将px值转换为Excel的磅值 (大约是0.75倍)
      excelCell.font.size = Math.round(parseInt(fontSize) * 0.75);
    }
  },
  
  /**
   * 在子元素中查找字体大小
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @returns {String} - 字体大小
   */
  findFontSizeInChildren(htmlCell) {
    if (!htmlCell) {
      return '';
    }
    
    let fontSize = '';
    const styledElements = htmlCell.querySelectorAll('*[style]');
    
    for (let i = 0; i < styledElements.length; i++) {
      const elem = styledElements[i];
      if (elem.style.fontSize) {
        fontSize = elem.style.fontSize;
        break;
      }
    }
    
    return fontSize;
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
    const colors = this.collectColorInfo(htmlCell);
    const textColor = colors.textColor;
    const backgroundColor = colors.backgroundColor;
    
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
   * 收集单元格内元素的颜色信息
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @returns {Object} - 颜色信息对象
   */
  collectColorInfo(htmlCell) {
    if (!htmlCell) {
      return { textColor: '', backgroundColor: '' };
    }
    
    let textColor = '';
    let backgroundColor = '';
    
    const styledElements = htmlCell.querySelectorAll('*[style]');
    Array.from(styledElements).forEach(function(elem) {
      const style = elem.style;
      if (style.color) textColor = style.color;
      if (style.backgroundColor) backgroundColor = style.backgroundColor;
    });
    
    return { textColor, backgroundColor };
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
   * 计算内容的显示宽度 (主方法)
   * @param {String} content - 内容
   * @returns {Number} - 估算的像素宽度
   */
  calculateContentWidth(content) {
    if (!content) return 0;
    
    // 处理多行内容情况
    if (content.includes('\n')) {
      return this.calculateMultilineWidth(content);
    }
    
    // 计算单行内容宽度
    return this.calculateSingleLineWidth(content);
  },
  
  /**
   * 计算多行内容宽度
   * @param {String} content - 多行内容
   * @returns {Number} - 估算的像素宽度
   */
  calculateMultilineWidth(content) {
    const lines = content.split('\n');
    let maxLineWidth = 0;
    
    for (let i = 0; i < lines.length; i++) {
      const lineWidth = this.calculateSingleLineWidth(lines[i]);
      maxLineWidth = Math.max(maxLineWidth, lineWidth);
    }
    
    return maxLineWidth;
  },
  
  /**
   * 计算单行内容宽度
   * @param {String} line - 单行内容
   * @returns {Number} - 估算的像素宽度
   */
  calculateSingleLineWidth(line) {
    // 基础字符宽度（像素）
    const charWidths = {
      default: 7,  // 默认ASCII字符宽度
      wide: 12,    // 宽字符（如中文）宽度
      narrow: 4    // 窄字符宽度
    };
    
    let totalWidth = 0;
    
    // 统计不同类型字符
    for (let i = 0; i < line.length; i++) {
      const charWidth = this.getCharacterWidth(line.charAt(i), charWidths);
      totalWidth += charWidth;
    }
    
    return totalWidth;
  },
  
  /**
   * 获取单个字符的宽度
   * @param {String} char - 字符
   * @param {Object} charWidths - 字符宽度配置
   * @returns {Number} - 字符宽度
   */
  getCharacterWidth(char, charWidths) {
    const code = char.charCodeAt(0);
    
    // 中文字符和其他全角字符
    if (this.isWideCharacter(code)) {
      return charWidths.wide;
    } 
    // 窄字符
    else if (this.isNarrowCharacter(char)) {
      return charWidths.narrow;
    } 
    // 默认ASCII字符
    else {
      return charWidths.default;
    }
  },
  
  /**
   * 判断是否为宽字符（如中文、日文等）
   * @param {Number} code - 字符码点
   * @returns {Boolean} - 是否为宽字符
   */
  isWideCharacter(code) {
    return (code >= 0x4E00 && code <= 0x9FFF) || 
           (code >= 0x3000 && code <= 0x303F) || 
           (code >= 0xFF00 && code <= 0xFFEF);
  },
  
  /**
   * 判断是否为窄字符
   * @param {String} char - 字符
   * @returns {Boolean} - 是否为窄字符
   */
  isNarrowCharacter(char) {
    const narrowChars = ['i', 'l', 'I', '.', ',', ':', ';'];
    return narrowChars.indexOf(char) !== -1;
  }
};
