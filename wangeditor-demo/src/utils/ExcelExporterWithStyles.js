/**
 * Excel导出工具类（带样式版本）
 */
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

export default {
  /**
   * 将表格导出为Excel文件（支持样式和合并单元格）
   * @param {HTMLElement} editorElement - 编辑器元素
   * @param {String} fileName - 导出的文件名(可选)
   * @returns {Boolean} - 是否成功导出
   */
  exportTableToExcel(editorElement, fileName = `表格导出_${new Date().toISOString().slice(0, 10)}`) {
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
      const workbook = XLSX.utils.book_new();
      
      // 处理每一个表格
      Array.from(tables).forEach((table, index) => {
        // 从表格生成工作表数据和样式信息
        const worksheet = this.createStyledWorksheet(table);
        
        // 添加工作表到工作簿
        const sheetName = `表格${index + 1}`;
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
      });
      
      // 保存文件 - 使用自定义write函数支持样式
      const wbout = this.writeWorkbookWithStyles(workbook);
      
      // 创建Blob
      const blob = new Blob([wbout], { type: 'application/octet-stream' });
      
      // 保存文件
      saveAs(blob, `${fileName}.xlsx`);
      
      return true;
    } catch (error) {
      console.error('导出Excel失败:', error);
      return false;
    }
  },
  
  /**
   * 创建带样式的工作表
   * @param {HTMLTableElement} table - 表格元素
   * @returns {Object} - 工作表对象
   */
  createStyledWorksheet(table) {
    const data = [];
    const merges = [];
    const styles = {};
    const skipCells = {};
    
    const rows = table.querySelectorAll('tr');
    
    Array.from(rows).forEach((row, rowIndex) => {
      const rowData = [];
      const cells = row.querySelectorAll('th, td');
      let colIndex = 0;
      
      Array.from(cells).forEach(cell => {
        // 检查是否需要跳过当前单元格
        while (skipCells[`${rowIndex},${colIndex}`]) {
          rowData.push('');
          colIndex++;
        }
        
        // 获取单元格文本内容
        const text = cell.textContent.trim();
        
        // 提取样式信息
        const cellStyle = this.extractCellStyle(cell);
        styles[`${rowIndex},${colIndex}`] = cellStyle;
        
        // 处理colspan和rowspan
        const colspan = parseInt(cell.getAttribute('colspan')) || 1;
        const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
        
        // 当前单元格的内容
        rowData.push({
          v: text,
          s: cellStyle
        });
        
        // 添加合并单元格信息
        if (colspan > 1 || rowspan > 1) {
          merges.push({
            s: { r: rowIndex, c: colIndex },
            e: { r: rowIndex + rowspan - 1, c: colIndex + colspan - 1 }
          });
          
          // 标记被合并的单元格
          for (let i = 0; i < rowspan; i++) {
            for (let j = 0; j < colspan; j++) {
              if (i !== 0 || j !== 0) {
                skipCells[`${rowIndex + i},${colIndex + j}`] = true;
              }
            }
          }
        }
        
        colIndex++;
        
        // 填充colspan的额外空单元格
        for (let i = 1; i < colspan; i++) {
          rowData.push('');
          colIndex++;
        }
      });
      
      data.push(rowData);
    });
    
    // 创建工作表
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    
    // 添加合并单元格信息
    if (merges.length > 0) {
      worksheet['!merges'] = merges;
    }
    
    // 设置列宽
    const colWidths = this.calculateColumnWidths(table);
    worksheet['!cols'] = colWidths;
    
    return worksheet;
  },
  
  /**
   * 从HTML单元格提取样式信息
   * @param {HTMLElement} cell - 单元格元素
   * @returns {Object} - 样式对象
   */
  extractCellStyle(cell) {
    const style = {};
    const computedStyle = window.getComputedStyle(cell);
    const inlineStyle = cell.style;
    
    // 填充色
    if (inlineStyle.backgroundColor || computedStyle.backgroundColor !== 'rgba(0, 0, 0, 0)') {
      const bgColor = inlineStyle.backgroundColor || computedStyle.backgroundColor;
      style.fill = {
        patternType: 'solid',
        fgColor: { rgb: this.colorToRGB(bgColor) }
      };
    }
    
    // 字体样式
    const font = {};
    
    // 字体颜色
    if (inlineStyle.color || computedStyle.color !== 'rgb(0, 0, 0)') {
      const color = inlineStyle.color || computedStyle.color;
      font.color = { rgb: this.colorToRGB(color) };
    }
    
    // 字体粗细
    if (inlineStyle.fontWeight === 'bold' || computedStyle.fontWeight === 'bold' || parseInt(computedStyle.fontWeight) >= 700) {
      font.bold = true;
    }
    
    // 字体大小
    if (inlineStyle.fontSize || computedStyle.fontSize) {
      font.sz = parseInt(inlineStyle.fontSize || computedStyle.fontSize);
    }
    
    // 斜体
    if (inlineStyle.fontStyle === 'italic' || computedStyle.fontStyle === 'italic') {
      font.italic = true;
    }
    
    // 下划线
    if (inlineStyle.textDecoration?.includes('underline') || computedStyle.textDecoration?.includes('underline')) {
      font.underline = true;
    }
    
    if (Object.keys(font).length > 0) {
      style.font = font;
    }
    
    // 对齐方式
    const alignment = {};
    
    // 水平对齐
    if (inlineStyle.textAlign || computedStyle.textAlign !== 'start') {
      const align = inlineStyle.textAlign || computedStyle.textAlign;
      alignment.horizontal = align === 'start' ? 'left' : align;
    }
    
    // 垂直对齐
    if (inlineStyle.verticalAlign || computedStyle.verticalAlign !== 'baseline') {
      const valign = inlineStyle.verticalAlign || computedStyle.verticalAlign;
      if (valign === 'middle') alignment.vertical = 'center';
      else if (valign === 'top') alignment.vertical = 'top';
      else if (valign === 'bottom') alignment.vertical = 'bottom';
    }
    
    if (Object.keys(alignment).length > 0) {
      style.alignment = alignment;
    }
    
    // 边框
    style.border = {
      top: { style: 'thin', color: { auto: 1 } },
      right: { style: 'thin', color: { auto: 1 } },
      bottom: { style: 'thin', color: { auto: 1 } },
      left: { style: 'thin', color: { auto: 1 } }
    };
    
    return style;
  },
  
  /**
   * 将颜色值转换为RGB格式
   * @param {String} color - 颜色值（支持rgb, rgba, hex）
   * @returns {String} - RGB格式颜色值
   */
  colorToRGB(color) {
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
   * 计算列宽
   * @param {HTMLTableElement} table - 表格元素
   * @returns {Array} - 列宽信息数组
   */
  calculateColumnWidths(table) {
    const colWidths = [];
    const rows = table.querySelectorAll('tr');
    
    // 获取最大列数
    let maxCols = 0;
    Array.from(rows).forEach(row => {
      const cells = row.querySelectorAll('th, td');
      let colCount = 0;
      Array.from(cells).forEach(cell => {
        colCount += parseInt(cell.getAttribute('colspan')) || 1;
      });
      maxCols = Math.max(maxCols, colCount);
    });
    
    // 设置默认列宽
    for (let i = 0; i < maxCols; i++) {
      colWidths.push({ wch: 15 }); // 默认列宽15
    }
    
    return colWidths;
  },
  
  /**
   * 写入带样式的工作簿
   * @param {Object} workbook - 工作簿对象
   * @returns {ArrayBuffer} - 二进制数据
   */
  writeWorkbookWithStyles(workbook) {
    // 创建自定义样式表
    const styles = this.createStylesheet();
    workbook.SSF = styles.numFmts;
    workbook.Styles = styles;
    
    const wopts = {
      bookType: 'xlsx',
      bookSST: false,
      type: 'buffer',
      cellStyles: true
    };
    
    return XLSX.write(workbook, wopts);
  },
  
  /**
   * 创建样式表
   * @returns {Object} - 样式表对象
   */
  createStylesheet() {
    return {
      numFmts: [],
      fonts: [
        { name: '宋体', sz: 11, color: { rgb: '000000' } },
        { name: 'Arial', sz: 10, color: { rgb: '000000' }, bold: true }
      ],
      fills: [
        { patternType: 'none' },
        { patternType: 'gray125' },
        { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } }
      ],
      borders: [
        {},
        {
          left: { style: 'thin' },
          right: { style: 'thin' },
          top: { style: 'thin' },
          bottom: { style: 'thin' }
        }
      ],
      cellStyleXfs: [
        { fontId: 0, fillId: 0, borderId: 0 }
      ],
      cellXfs: [
        { fontId: 0, fillId: 0, borderId: 0 }
      ]
    };
  }
}