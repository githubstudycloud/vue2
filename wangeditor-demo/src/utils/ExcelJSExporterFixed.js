/**
 * Excel导出工具类（使用ExcelJS - 数据修复版）
 */
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default {
  /**
   * 将表格导出为Excel文件（支持完整样式和合并单元格）
   * @param {HTMLElement} editorElement - 编辑器元素
   * @param {String} fileName - 导出的文件名(可选)
   * @returns {Boolean} - 是否成功导出
   */
  async exportTableToExcel(editorElement, fileName = `表格导出_${new Date().toISOString().slice(0, 10)}`) {
    if (!editorElement) {
      console.error('编辑器元素不存在');
      return false;
    }
    
    // 获取编辑器中的表格
    const tables = editorElement.querySelectorAll('table');
    if (!tables || tables.length === 0) {
      console.error('未找到可导出的表格');
      return false;
    }

    try {
      // 创建工作簿
      const workbook = new ExcelJS.Workbook();
      
      // 处理每一个表格
      Array.from(tables).forEach((table, index) => {
        // 跳过嵌套表格（父表格中的子表格）
        if (this.isNestedTable(table)) {
          return;
        }
        
        // 添加工作表
        const worksheet = workbook.addWorksheet(`表格${index + 1}`);
        
        // 处理表格数据
        this.processTable(table, worksheet);
      });
      
      // 生成Excel文件
      const buffer = await workbook.xlsx.writeBuffer();
      
      // 保存文件
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `${fileName}.xlsx`);
      
      return true;
    } catch (error) {
      console.error('导出Excel失败:', error);
      return false;
    }
  },
  
  /**
   * 检查是否是嵌套表格
   * @param {HTMLTableElement} table - 表格元素
   * @returns {Boolean} - 是否是嵌套表格
   */
  isNestedTable(table) {
    let parent = table.parentElement;
    while (parent) {
      if (parent.tagName === 'TABLE') {
        return true;
      }
      parent = parent.parentElement;
    }
    return false;
  },
  
  /**
   * 处理表格数据
   * @param {HTMLTableElement} table - 表格元素
   * @param {Object} worksheet - 工作表对象
   */
  processTable(table, worksheet) {
    const rows = table.querySelectorAll('tr');
    const skipCells = {};
    
    Array.from(rows).forEach((row, rowIndex) => {
      const cells = row.querySelectorAll('th, td');
      let colIndex = 1; // ExcelJS中列索引从1开始
      
      Array.from(cells).forEach(cell => {
        // 检查是否需要跳过当前单元格
        while (skipCells[`${rowIndex + 1},${colIndex}`]) {
          colIndex++;
        }
        
        // 处理包含嵌套表格的单元格
        const nestedTable = cell.querySelector('table');
        let text = '';
        
        if (nestedTable) {
          // 将嵌套表格转换为文本内容
          text = this.convertNestedTableToText(nestedTable);
        } else {
          // 获取单元格文本内容
          text = cell.textContent.trim();
        }
        
        // 设置单元格值
        const excelCell = worksheet.getCell(rowIndex + 1, colIndex);
        excelCell.value = text;
        
        // 应用样式
        this.applyCellStyle(cell, excelCell);
        
        // 处理合并单元格
        const colspan = parseInt(cell.getAttribute('colspan')) || 1;
        const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
        
        if (colspan > 1 || rowspan > 1) {
          worksheet.mergeCells(
            rowIndex + 1,
            colIndex,
            rowIndex + rowspan,
            colIndex + colspan - 1
          );
          
          // 标记被合并的单元格
          for (let i = 0; i < rowspan; i++) {
            for (let j = 0; j < colspan; j++) {
              if (i !== 0 || j !== 0) {
                skipCells[`${rowIndex + 1 + i},${colIndex + j}`] = true;
              }
            }
          }
        }
        
        colIndex += colspan;
      });
    });
    
    // 设置列宽
    this.setColumnWidths(worksheet);
    
    // 设置行高
    this.setRowHeights(worksheet);
  },
  
  /**
   * 将嵌套表格转换为文本内容
   * @param {HTMLTableElement} table - 嵌套表格元素
   * @returns {String} - 文本内容
   */
  convertNestedTableToText(table) {
    let text = '';
    const rows = table.querySelectorAll('tr');
    
    Array.from(rows).forEach((row, rowIndex) => {
      const cells = row.querySelectorAll('th, td');
      let rowText = '';
      
      Array.from(cells).forEach((cell, cellIndex) => {
        if (cellIndex > 0) {
          rowText += ' | ';
        }
        rowText += cell.textContent.trim();
      });
      
      if (rowIndex > 0) {
        text += '\n';
      }
      text += rowText;
    });
    
    return text;
  },
  
  /**
   * 应用单元格样式
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   */
  applyCellStyle(htmlCell, excelCell) {
    // 默认字体样式
    excelCell.font = {
      name: '宋体', // 默认字体
      size: 11,
      bold: false,
      italic: false,
      underline: false,
      color: { argb: 'FF000000' }
    };
    
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
    
    // 检查是否有带样式的子元素
    let textColor = '';
    let backgroundColor = '';
    let fontSize = '';
    
    const styledElements = htmlCell.querySelectorAll('*[style]');
    Array.from(styledElements).forEach(elem => {
      const style = elem.style;
      // 收集颜色信息
      if (style.color) textColor = style.color;
      if (style.backgroundColor) backgroundColor = style.backgroundColor;
      if (style.fontSize) fontSize = style.fontSize;
      
      // 检查文本装饰
      if (style.fontWeight === 'bold' || parseInt(style.fontWeight) >= 700) {
        excelCell.font.bold = true;
      }
      if (style.fontStyle === 'italic') {
        excelCell.font.italic = true;
      }
      if (style.textDecoration?.includes('underline')) {
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
    
    // 字体颜色
    if (textColor) {
      const argb = this.colorToARGB(textColor);
      excelCell.font.color = { argb: argb };
    } else if (inlineStyle.color) {
      const argb = this.colorToARGB(inlineStyle.color);
      excelCell.font.color = { argb: argb };
    } else if (computedStyle.color && computedStyle.color !== 'rgb(0, 0, 0)') {
      const argb = this.colorToARGB(computedStyle.color);
      excelCell.font.color = { argb: argb };
    }
    
    // 字体粗细
    if (!excelCell.font.bold && (inlineStyle.fontWeight === 'bold' || computedStyle.fontWeight === 'bold' || parseInt(computedStyle.fontWeight) >= 700)) {
      excelCell.font.bold = true;
    }
    
    // 字体斜体
    if (!excelCell.font.italic && (inlineStyle.fontStyle === 'italic' || computedStyle.fontStyle === 'italic')) {
      excelCell.font.italic = true;
    }
    
    // 下划线
    if (!excelCell.font.underline && (inlineStyle.textDecoration?.includes('underline') || computedStyle.textDecoration?.includes('underline'))) {
      excelCell.font.underline = true;
    }
    
    // 字体大小
    if (fontSize) {
      excelCell.font.size = Math.round(parseInt(fontSize) * 0.75);
    } else if (inlineStyle.fontSize) {
      excelCell.font.size = Math.round(parseInt(inlineStyle.fontSize) * 0.75);
    } else if (computedStyle.fontSize) {
      excelCell.font.size = Math.round(parseInt(computedStyle.fontSize) * 0.75);
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
    } else if (computedStyle.backgroundColor && computedStyle.backgroundColor !== 'rgba(0, 0, 0, 0)') {
      const argb = this.colorToARGB(computedStyle.backgroundColor);
      excelCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: argb }
      };
    }
    
    // 对齐方式
    const alignment = {
      wrapText: true, // 自动换行
      vertical: 'middle' // 默认垂直居中
    };
    
    // 水平对齐
    const textAlign = inlineStyle.textAlign || computedStyle.textAlign;
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
    const verticalAlign = inlineStyle.verticalAlign || computedStyle.verticalAlign;
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
    
    // 完整的边框样式
    excelCell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };
  },
  
  /**
   * 映射网页字体到Excel支持的字体
   * @param {String} webFont - 网页字体名称
   * @returns {String} - Excel支持的字体名称
   */
  mapWebFontToExcel(webFont) {
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
        hex = hex.split('').map(c => c + c).join('');
      }
      return 'FF' + hex.toUpperCase();
    }
    
    // 处理rgb/rgba格式
    const rgbaMatch = color.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([0-9.]+))?\)$/);
    if (rgbaMatch) {
      const [, r, g, b, a] = rgbaMatch;
      let alpha = 'FF';
      if (a !== undefined) {
        alpha = Math.round(parseFloat(a) * 255).toString(16).toUpperCase().padStart(2, '0');
      }
      return alpha + this.numberToHex(r) + this.numberToHex(g) + this.numberToHex(b);
    }
    
    // 处理颜色名称映射
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
    if (colorMap[colorName]) {
      return colorMap[colorName];
    }
    
    return 'FF000000'; // 默认黑色
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
   * 设置列宽
   * @param {Object} worksheet - 工作表对象
   */
  setColumnWidths(worksheet) {
    // 获取所有列
    const columns = worksheet.columns;
    
    columns.forEach((column, index) => {
      let maxLength = 12; // 最小列宽
      
      // 遍历该列的所有单元格，找出最长的内容
      column.eachCell({ includeEmpty: true }, cell => {
        const cellValue = cell.value;
        if (cellValue) {
          let length = cellValue.toString().length;
          
          // 为中文字符增加额外宽度
          const chineseChars = cellValue.toString().match(/[\u4e00-\u9fa5]/g);
          if (chineseChars) {
            length += chineseChars.length * 0.8;
          }
          
          if (length > maxLength) {
            maxLength = length;
          }
        }
      });
      
      // 设置列宽，添加一些余量
      column.width = Math.min(maxLength * 1.2, 60); // 最大宽度限制为60
    });
  },
  
  /**
   * 设置行高
   * @param {Object} worksheet - 工作表对象
   */
  setRowHeights(worksheet) {
    // 遍历所有行
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      let maxHeight = 20; // 默认行高
      let pixelHeight = 0;
      
      // 检查内容高度
      row.eachCell({ includeEmpty: false }, cell => {
        const cellValue = cell.value;
        if (cellValue) {
          // 如果内容包含换行符，增加行高
          const lines = cellValue.toString().split('\n').length;
          const cellHeight = lines * 20;
          maxHeight = Math.max(maxHeight, cellHeight);
        }
      });
      
      row.height = maxHeight;
    });
  }
}