/**
 * Excel导出工具类（完整版，支持完整样式）
 * 使用sheetjs-style来支持样式导出
 */
import XLSX from 'xlsx';
import { saveAs } from 'file-saver';

// 扩展XLSX以支持样式
if (typeof window !== 'undefined') {
  window.XLSX = XLSX;
  // 导入xlsx-style-vite以支持样式
  import('xlsx-style-vite').then(XLSXStyle => {
    Object.assign(XLSX, XLSXStyle);
  });
}

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
      // 等待xlsx-style-vite加载完成
      await new Promise(resolve => setTimeout(resolve, 100));
      
      // 创建工作簿
      const workbook = XLSX.utils.book_new();
      
      // 处理每一个表格
      Array.from(tables).forEach((table, index) => {
        // 从表格生成工作表数据和样式信息
        const worksheet = XLSX.utils.table_to_sheet(table);
        
        // 解析样式并应用到工作表
        this.applyStyles(table, worksheet);
        
        // 处理合并单元格
        const merges = this.findMergedCells(table);
        if (merges.length > 0) {
          worksheet['!merges'] = merges;
        }
        
        // 设置列宽
        const colWidths = this.calculateColumnWidths(table);
        worksheet['!cols'] = colWidths;
        
        // 添加工作表到工作簿
        const sheetName = `表格${index + 1}`;
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
      });
      
      // 使用xlsx-style写入带样式的Excel
      const wbout = XLSX.write(workbook, {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary',
        cellStyles: true
      });
      
      // 字符串转ArrayBuffer
      function s2ab(s) {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) {
          view[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf;
      }
      
      // 保存文件
      saveAs(new Blob([s2ab(wbout)], { type: 'application/octet-stream' }), `${fileName}.xlsx`);
      
      return true;
    } catch (error) {
      console.error('导出Excel失败:', error);
      return false;
    }
  },
  
  /**
   * 应用样式到工作表
   * @param {HTMLTableElement} table - 表格元素
   * @param {Object} worksheet - 工作表对象
   */
  applyStyles(table, worksheet) {
    const rows = table.querySelectorAll('tr');
    let skipCells = {};
    
    Array.from(rows).forEach((row, rowIndex) => {
      const cells = row.querySelectorAll('th, td');
      let colIndex = 0;
      
      Array.from(cells).forEach(cell => {
        // 检查是否需要跳过当前单元格（被合并的单元格）
        while (skipCells[`${rowIndex},${colIndex}`]) {
          colIndex++;
        }
        
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        
        if (worksheet[cellAddress]) {
          const style = this.extractCellStyle(cell);
          worksheet[cellAddress].s = style;
        }
        
        // 处理合并单元格的跳过
        const colspan = parseInt(cell.getAttribute('colspan')) || 1;
        const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
        
        // 标记被合并的单元格
        if (colspan > 1 || rowspan > 1) {
          for (let i = 0; i < rowspan; i++) {
            for (let j = 0; j < colspan; j++) {
              if (i !== 0 || j !== 0) {
                skipCells[`${rowIndex + i},${colIndex + j}`] = true;
              }
            }
          }
        }
        
        colIndex += colspan;
      });
    });
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
    const bgColor = inlineStyle.backgroundColor || computedStyle.backgroundColor;
    if (bgColor && bgColor !== 'rgba(0, 0, 0, 0)' && bgColor !== 'transparent') {
      style.fill = {
        patternType: 'solid',
        fgColor: { rgb: this.colorToRGB(bgColor) }
      };
    }
    
    // 字体样式
    const font = {};
    
    // 字体颜色
    const color = inlineStyle.color || computedStyle.color;
    if (color && color !== 'rgb(0, 0, 0)') {
      font.color = { rgb: this.colorToRGB(color) };
    }
    
    // 字体粗细
    const fontWeight = inlineStyle.fontWeight || computedStyle.fontWeight;
    if (fontWeight === 'bold' || parseInt(fontWeight) >= 700) {
      font.bold = true;
    }
    
    // 字体大小
    const fontSize = parseInt(inlineStyle.fontSize || computedStyle.fontSize);
    if (fontSize) {
      font.sz = Math.round(fontSize * 0.75); // Excel的字体大小单位是pt，而不是px
    }
    
    // 斜体
    const fontStyle = inlineStyle.fontStyle || computedStyle.fontStyle;
    if (fontStyle === 'italic') {
      font.italic = true;
    }
    
    // 下划线
    const textDecoration = inlineStyle.textDecoration || computedStyle.textDecoration;
    if (textDecoration && textDecoration.includes('underline')) {
      font.underline = true;
    }
    
    // 字体名称
    const fontFamily = inlineStyle.fontFamily || computedStyle.fontFamily;
    if (fontFamily) {
      font.name = fontFamily.split(',')[0].replace(/['"]/g, '').trim();
    }
    
    if (Object.keys(font).length > 0) {
      style.font = font;
    }
    
    // 对齐方式
    const alignment = {};
    
    // 水平对齐
    const textAlign = inlineStyle.textAlign || computedStyle.textAlign;
    if (textAlign) {
      alignment.horizontal = textAlign === 'start' ? 'left' : textAlign;
    }
    
    // 垂直对齐
    const verticalAlign = inlineStyle.verticalAlign || computedStyle.verticalAlign;
    if (verticalAlign === 'middle') {
      alignment.vertical = 'center';
    } else if (verticalAlign === 'top' || verticalAlign === 'baseline') {
      alignment.vertical = 'top';
    } else if (verticalAlign === 'bottom') {
      alignment.vertical = 'bottom';
    }
    
    if (Object.keys(alignment).length > 0) {
      style.alignment = alignment;
    }
    
    // 边框
    style.border = {
      top: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } }
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
      return 'FF' + hex.toUpperCase(); // FF表示完全不透明
    }
    
    // 处理rgb/rgba格式
    const match = color.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/);
    if (match) {
      const [, r, g, b, a] = match;
      let alpha = 'FF';
      if (a !== undefined) {
        alpha = Math.round(parseFloat(a) * 255).toString(16).toUpperCase().padStart(2, '0');
      }
      return alpha + this.numberToHex(r) + this.numberToHex(g) + this.numberToHex(b);
    }
    
    // 处理颜色名称
    const colorMap = {
      'black': 'FF000000',
      'white': 'FFFFFFFF',
      'red': 'FFFF0000',
      'green': 'FF00FF00',
      'blue': 'FF0000FF',
      // ... 可以添加更多颜色映射
    };
    
    return colorMap[color.toLowerCase()] || 'FF000000'; // 默认黑色
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
          
          // 标记被合并的单元格
          for (let i = 0; i < rowspan; i++) {
            for (let j = 0; j < colspan; j++) {
              if (i !== 0 || j !== 0) {
                skipCells[`${rowIndex + i},${colIndex + j}`] = true;
              }
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
  }
}