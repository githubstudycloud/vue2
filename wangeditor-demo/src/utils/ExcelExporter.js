/**
 * Excel导出工具类
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
        // 从表格生成工作表数据
        const tableData = this.parseTableToArray(table);
        
        // 创建工作表
        const worksheet = XLSX.utils.aoa_to_sheet(tableData);
        
        // 处理合并单元格
        const merges = this.findMergedCells(table);
        if (merges.length > 0) {
          worksheet['!merges'] = merges;
        }
        
        // 处理列宽
        const colWidths = this.calculateColumnWidths(table);
        worksheet['!cols'] = colWidths;
        
        // 处理样式（如果支持）
        this.applyStyles(table, worksheet);
        
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
   * 将HTML表格解析为二维数组
   * @param {HTMLTableElement} table - 表格元素
   * @returns {Array} - 二维数组形式的表格数据
   */
  parseTableToArray(table) {
    const rows = table.querySelectorAll('tr');
    const data = [];
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
    
    return data;
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
    Array.from(cells).forEach((cell) => {
      // 根据内容长度估算列宽
      const text = cell.textContent.trim();
      const width = Math.max(text.length * 2, 10);
      colWidths.push({ width });
    });
    
    return colWidths;
  },
  
  /**
   * 应用样式（注：由于xlsx不支持所有样式，此功能有限）
   * @param {HTMLTableElement} table - 表格元素
   * @param {Object} worksheet - 工作表对象
   */
  applyStyles(table, worksheet) {
    // 由于xlsx的标准版本对样式支持有限，这里暂不实现
    // 如果需要完整样式支持，建议使用xlsx-style或其他专门的库
  }
}