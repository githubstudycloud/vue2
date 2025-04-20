/**
 * Excel导出工具类（增强版，支持更多样式）
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
        const { data, styleMap } = this.parseTableToDataWithStyles(table);
        
        // 创建工作表
        const worksheet = XLSX.utils.aoa_to_sheet(data);
        
        // 处理合并单元格
        const merges = this.findMergedCells(table);
        if (merges.length > 0) {
          worksheet['!merges'] = merges;
        }
        
        // 处理列宽
        const colWidths = this.calculateColumnWidths(table);
        worksheet['!cols'] = colWidths;
        
        // 应用样式
        this.applyStylesToWorksheet(worksheet, styleMap);
        
        // 添加工作表到工作簿
        XLSX.utils.book_append_sheet(workbook, worksheet, `表格${index + 1}`);
      });
      
      // 转为二进制数据
      const excelBuffer = XLSX.write(workbook, { 
        bookType: 'xlsx', 
        type: 'array',
        cellStyles: true 
      });
      
      // 创建Blob
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      // 保存文件
      saveAs(blob, `${fileName}.xlsx`);
      
      return true;
    } catch (error) {
      console.error('导出Excel失败:', error);
      return false;
    }
  },
  
  /**
   * 将HTML表格解析为二维数组，并提取样式信息
   * @param {HTMLTableElement} table - 表格元素
   * @returns {Object} - 包含数据和样式映射的对象
   */
  parseTableToDataWithStyles(table) {
    const rows = table.querySelectorAll('tr');
    const data = [];
    const styleMap = {};
    const skipCells = {}; // 记录需要跳过的单元格位置
    
    Array.from(rows).forEach((row, rowIndex) => {
      const rowData = [];
      // 获取所有单元格（th和td）
      const cells = row.querySelectorAll('th, td');
      let colIndex = 0;
      
      Array.from(cells).forEach(cell => {
        // 检查是否需要跳过当前单元格
        while (skipCells[`${rowIndex},${colIndex}`]) {
          rowData.push(''); // 填充空值
          colIndex++;
        }
        
        // 获取单元格文本内容
        const text = cell.textContent.trim();
        
        // 提取样式信息
        const cellStyle = this.extractCellStyle(cell);
        if (Object.keys(cellStyle).length > 0) {
          styleMap[`${rowIndex},${colIndex}`] = cellStyle;
        }
        
        // 处理colspan
        const colspan = parseInt(cell.getAttribute('colspan')) || 1;
        // 处理rowspan
        const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
        
        // 当前单元格的内容
        rowData.push(text);
        colIndex++;
        
        // 填充colspan的额外空单元格
        for (let i = 1; i < colspan; i++) {
          rowData.push('');
          colIndex++;
        }
        
        // 标记rowspan影响的单元格
        for (let i = 1; i < rowspan; i++) {
          for (let j = 0; j < colspan; j++) {
            skipCells[`${rowIndex + i},${colIndex - colspan + j}`] = true;
          }
        }
      });
      
      data.push(rowData);
    });
    
    return { data, styleMap };
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
    
    // 背景色
    if (inlineStyle.backgroundColor || computedStyle.backgroundColor !== 'rgba(0, 0, 0, 0)') {
      style.fill = {
        fgColor: { rgb: this.colorToARGB(inlineStyle.backgroundColor || computedStyle.backgroundColor) }
      };
    }
    
    // 字体颜色、大小、粗细
    const font = {};
    if (inlineStyle.color || computedStyle.color !== 'rgb(0, 0, 0)') {
      font.color = { rgb: this.colorToARGB(inlineStyle.color || computedStyle.color) };
    }
    if (inlineStyle.fontWeight === 'bold' || computedStyle.fontWeight === 'bold' || computedStyle.fontWeight >= 700) {
      font.bold = true;
    }
    if (inlineStyle.fontSize || computedStyle.fontSize) {
      font.sz = parseInt(inlineStyle.fontSize || computedStyle.fontSize);
    }
    
    if (Object.keys(font).length > 0) {
      style.font = font;
    }
    
    // 对齐方式
    const alignment = {};
    if (inlineStyle.textAlign || computedStyle.textAlign) {
      const align = inlineStyle.textAlign || computedStyle.textAlign;
      if (align === 'center') alignment.horizontal = 'center';
      else if (align === 'right') alignment.horizontal = 'right';
    }
    if (inlineStyle.verticalAlign || computedStyle.verticalAlign) {
      const valign = inlineStyle.verticalAlign || computedStyle.verticalAlign;
      if (valign === 'middle') alignment.vertical = 'center';
    }
    
    if (Object.keys(alignment).length > 0) {
      style.alignment = alignment;
    }
    
    return style;
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
      color = color.slice(1);
      if (color.length === 3) {
        color = color.split('').map(c => c + c).join('');
      }
      return 'FF' + color.toUpperCase();
    }
    
    // 处理rgb/rgba格式
    const match = color.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+(?:\.\d+)?))?\)$/);
    if (match) {
      const [, r, g, b, a] = match;
      const alpha = a ? Math.round(parseFloat(a) * 255) : 255;
      return this.numberToHex(alpha) + this.numberToHex(r) + this.numberToHex(g) + this.numberToHex(b);
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
   * 应用样式到工作表
   * @param {Object} worksheet - 工作表对象
   * @param {Object} styleMap - 样式映射
   */
  applyStylesToWorksheet(worksheet, styleMap) {
    // 标准xlsx库不支持直接设置样式，这里仅作为参考
    // 在实际使用中，可能需要使用xlsx-style或其他支持样式的库
    worksheet['!styles'] = styleMap;
  },
  
  /**
   * 查找合并单元格
   * @param {HTMLTableElement} table - 表格元素
   * @returns {Array} - 合并单元格信息数组
   */
  findMergedCells(table) {
    const merges = [];
    const rows = table.querySelectorAll('tr');
    const skipCells = {};
    
    Array.from(rows).forEach((row, rowIndex) => {
      const cells = row.querySelectorAll('th, td');
      let colIndex = 0;
      
      Array.from(cells).forEach(cell => {
        // 跳过被合并的单元格
        while (skipCells[`${rowIndex},${colIndex}`]) {
          colIndex++;
        }
        
        const colspan = parseInt(cell.getAttribute('colspan')) || 1;
        const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
        
        // 添加合并单元格信息
        if (colspan > 1 || rowspan > 1) {
          merges.push({
            s: { r: rowIndex, c: colIndex },
            e: { r: rowIndex + rowspan - 1, c: colIndex + colspan - 1 }
          });
        }
        
        // 标记被合并的单元格
        for (let i = 0; i < rowspan; i++) {
          for (let j = 0; j < colspan; j++) {
            if (i !== 0 || j !== 0) {
              skipCells[`${rowIndex + i},${colIndex + j}`] = true;
            }
          }
        }
        
        colIndex += colspan;
      });
    });
    
    return merges;
  },
  
  /**
   * 计算列宽
   * @param {HTMLTableElement} table - 表格元素
   * @returns {Array} - 列宽信息数组
   */
  calculateColumnWidths(table) {
    const colWidths = [];
    const firstRow = table.querySelector('tr');
    if (!firstRow) return colWidths;
    
    const cells = firstRow.querySelectorAll('th, td');
    let colIndex = 0;
    
    Array.from(cells).forEach((cell) => {
      const colspan = parseInt(cell.getAttribute('colspan')) || 1;
      const text = cell.textContent.trim();
      const width = Math.max(text.length * 1.2, 10);
      
      for (let i = 0; i < colspan; i++) {
        colWidths[colIndex + i] = { width };
      }
      
      colIndex += colspan;
    });
    
    return colWidths;
  }
}