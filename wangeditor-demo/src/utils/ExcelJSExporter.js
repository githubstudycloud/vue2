/**
 * Excel导出工具类（使用ExcelJS）
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
  },
  
  /**
   * 应用单元格样式
   * @param {HTMLElement} htmlCell - HTML单元格元素
   * @param {Object} excelCell - Excel单元格对象
   */
  applyCellStyle(htmlCell, excelCell) {
    const computedStyle = window.getComputedStyle(htmlCell);
    const inlineStyle = htmlCell.style;
    
    // 字体样式
    excelCell.font = {
      name: '宋体', // 默认字体
      size: 11,
      bold: false,
      italic: false,
      underline: false,
      color: { argb: 'FF000000' }
    };
    
    // 字体颜色
    const color = inlineStyle.color || computedStyle.color;
    if (color && color !== 'rgb(0, 0, 0)') {
      excelCell.font.color = { argb: this.colorToARGB(color) };
    }
    
    // 字体粗细
    if (inlineStyle.fontWeight === 'bold' || computedStyle.fontWeight === 'bold' || parseInt(computedStyle.fontWeight) >= 700) {
      excelCell.font.bold = true;
    }
    
    // 字体大小
    if (inlineStyle.fontSize || computedStyle.fontSize) {
      excelCell.font.size = Math.round(parseInt(inlineStyle.fontSize || computedStyle.fontSize) * 0.75);
    }
    
    // 斜体
    if (inlineStyle.fontStyle === 'italic' || computedStyle.fontStyle === 'italic') {
      excelCell.font.italic = true;
    }
    
    // 下划线
    if (inlineStyle.textDecoration?.includes('underline') || computedStyle.textDecoration?.includes('underline')) {
      excelCell.font.underline = true;
    }
    
    // 背景色
    const bgColor = inlineStyle.backgroundColor || computedStyle.backgroundColor;
    if (bgColor && bgColor !== 'rgba(0, 0, 0, 0)' && bgColor !== 'transparent') {
      excelCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: this.colorToARGB(bgColor) }
      };
    }
    
    // 对齐方式
    const alignment = {};
    
    // 水平对齐
    const textAlign = inlineStyle.textAlign || computedStyle.textAlign;
    if (textAlign === 'center') {
      alignment.horizontal = 'center';
    } else if (textAlign === 'right') {
      alignment.horizontal = 'right';
    } else if (textAlign === 'left' || textAlign === 'start') {
      alignment.horizontal = 'left';
    }
    
    // 垂直对齐
    const verticalAlign = inlineStyle.verticalAlign || computedStyle.verticalAlign;
    if (verticalAlign === 'middle') {
      alignment.vertical = 'middle';
    } else if (verticalAlign === 'top') {
      alignment.vertical = 'top';
    } else if (verticalAlign === 'bottom') {
      alignment.vertical = 'bottom';
    }
    
    if (Object.keys(alignment).length > 0) {
      excelCell.alignment = alignment;
    }
    
    // 边框
    excelCell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  },
  
  /**
   * 将颜色值转换为ARGB格式
   * @param {String} color - 颜色值（支持rgb, rgba, hex）
   * @returns {String} - ARGB格式颜色值
   */
  colorToARGB(color) {
    if (!color) return 'FF000000';
    
    // 处理hex格式
    if (color.startsWith('#')) {
      let hex = color.slice(1);
      if (hex.length === 3) {
        hex = hex.split('').map(c => c + c).join('');
      }
      return 'FF' + hex.toUpperCase();
    }
    
    // 处理rgb/rgba格式
    const match = color.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+(?:\.\d+)?))?\)$/);
    if (match) {
      const [, r, g, b, a] = match;
      let alpha = 'FF';
      if (a !== undefined) {
        alpha = Math.round(parseFloat(a) * 255).toString(16).toUpperCase().padStart(2, '0');
      }
      return alpha + this.numberToHex(r) + this.numberToHex(g) + this.numberToHex(b);
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
    // 自动设置列宽
    worksheet.columns.forEach((column, index) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, cell => {
        const length = cell.value ? cell.value.toString().length : 0;
        if (length > maxLength) {
          maxLength = length;
        }
      });
      column.width = maxLength < 10 ? 10 : maxLength + 2;
    });
  }
}