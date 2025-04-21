/**
 * Excel导出工具类（使用ExcelJS - 调试版）
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
        
        // 获取单元格文本内容
        const text = cell.textContent.trim();
        
        // 日志：输出单元格内容
        console.log(`处理单元格 [${rowIndex + 1}, ${colIndex}] 内容:`, text);
        
        // 设置单元格值
        const excelCell = worksheet.getCell(rowIndex + 1, colIndex);
        excelCell.value = text;
        
        // 应用样式
        this.applyCellStyle(cell, excelCell);
        
        // 处理合并单元格
        const colspan = parseInt(cell.getAttribute('colspan')) || 1;
        const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
        
        if (colspan > 1 || rowspan > 1) {
          console.log(`合并单元格 [${rowIndex + 1}, ${colIndex}] 范围: ${rowspan}行 x ${colspan}列`);
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
   * 应用单元格样式
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   */
  applyCellStyle(htmlCell, excelCell) {
    const computedStyle = window.getComputedStyle(htmlCell);
    const inlineStyle = htmlCell.style;
    
    // 日志：输出原始样式信息
    console.log('=== 样式信息 ===');
    console.log('内联样式 backgroundColor:', inlineStyle.backgroundColor);
    console.log('内联样式 color:', inlineStyle.color);
    console.log('计算样式 backgroundColor:', computedStyle.backgroundColor);
    console.log('计算样式 color:', computedStyle.color);
    console.log('计算样式 fontWeight:', computedStyle.fontWeight);
    console.log('计算样式 textAlign:', computedStyle.textAlign);
    console.log('===============');
    
    // 字体样式
    excelCell.font = {
      name: '宋体', // 默认字体
      size: 11,
      bold: false,
      italic: false,
      underline: false,
      color: { argb: 'FF000000' }
    };
    
    // 字体名称
    const fontFamily = inlineStyle.fontFamily || computedStyle.fontFamily;
    if (fontFamily) {
      excelCell.font.name = fontFamily.split(',')[0].replace(/['"]/g, '').trim();
      console.log('字体:', excelCell.font.name);
    }
    
    // 字体颜色
    const color = inlineStyle.color || computedStyle.color;
    if (color && color !== 'rgb(0, 0, 0)') {
      const argb = this.colorToARGB(color);
      excelCell.font.color = { argb: argb };
      console.log('字体颜色:', color, '-> ARGB:', argb);
    }
    
    // 字体粗细
    const fontWeight = inlineStyle.fontWeight || computedStyle.fontWeight;
    if (fontWeight === 'bold' || parseInt(fontWeight) >= 700) {
      excelCell.font.bold = true;
      console.log('字体粗细: 粗体');
    }
    
    // 字体大小
    const fontSize = inlineStyle.fontSize || computedStyle.fontSize;
    if (fontSize) {
      excelCell.font.size = Math.round(parseInt(fontSize) * 0.75);
      console.log('字体大小:', fontSize, '->', excelCell.font.size);
    }
    
    // 背景色
    const bgColor = inlineStyle.backgroundColor || computedStyle.backgroundColor;
    if (bgColor && bgColor !== 'rgba(0, 0, 0, 0)' && bgColor !== 'transparent') {
      const argb = this.colorToARGB(bgColor);
      excelCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: argb }
      };
      console.log('背景色:', bgColor, '-> ARGB:', argb);
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
      console.log('水平对齐:', textAlign, '->', alignment.horizontal);
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
      console.log('垂直对齐:', verticalAlign, '->', alignment.vertical);
    }
    
    excelCell.alignment = alignment;
    
    // 边框
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
    console.log('转换颜色:', color);
    
    if (!color) {
      console.log('颜色为空，返回默认黑色');
      return 'FF000000';
    }
    
    // 处理hex格式
    if (color.startsWith('#')) {
      let hex = color.slice(1);
      if (hex.length === 3) {
        hex = hex.split('').map(c => c + c).join('');
      }
      const result = 'FF' + hex.toUpperCase();
      console.log('hex格式:', color, '->', result);
      return result;
    }
    
    // 处理rgb/rgba格式
    const rgbaMatch = color.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([0-9.]+))?\)$/);
    if (rgbaMatch) {
      const [, r, g, b, a] = rgbaMatch;
      let alpha = 'FF';
      if (a !== undefined) {
        alpha = Math.round(parseFloat(a) * 255).toString(16).toUpperCase().padStart(2, '0');
      }
      const result = alpha + this.numberToHex(r) + this.numberToHex(g) + this.numberToHex(b);
      console.log('rgb/rgba格式:', color, '->', result);
      return result;
    }
    
    // 处理颜色名称映射
    const colorMap = {
      'black': 'FF000000',
      'white': 'FFFFFFFF',
      'red': 'FFFF0000',
      'green': 'FF00FF00',
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
      console.log('颜色名称:', color, '->', colorMap[colorName]);
      return colorMap[colorName];
    }
    
    console.log('无法识别的颜色格式，返回默认黑色');
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
      let maxLength = 10; // 最小列宽
      
      // 遍历该列的所有单元格，找出最长的内容
      column.eachCell({ includeEmpty: true }, cell => {
        const cellValue = cell.value;
        if (cellValue) {
          let length = cellValue.toString().length;
          
          // 为中文字符增加额外宽度
          const chineseChars = cellValue.toString().match(/[\u4e00-\u9fa5]/g);
          if (chineseChars) {
            length += chineseChars.length * 0.5;
          }
          
          if (length > maxLength) {
            maxLength = length;
          }
        }
      });
      
      // 设置列宽，添加一些余量
      column.width = Math.min(maxLength * 1.2, 50); // 最大宽度限制为50
      console.log(`列 ${index + 1} 宽度设置为:`, column.width);
    });
  },
  
  /**
   * 设置行高
   * @param {Object} worksheet - 工作表对象
   */
  setRowHeights(worksheet) {
    // 遍历所有行
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      let maxHeight = 18; // 默认行高
      
      // 检查行中的内容
      row.eachCell({ includeEmpty: false }, cell => {
        const cellValue = cell.value;
        if (cellValue) {
          // 如果内容包含换行符，增加行高
          const lines = cellValue.toString().split('\n').length;
          const cellHeight = lines * 18;
          maxHeight = Math.max(maxHeight, cellHeight);
        }
      });
      
      row.height = maxHeight;
      console.log(`行 ${rowNumber} 高度设置为:`, row.height);
    });
  }
}