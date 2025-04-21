/**
 * Excel导出工具类（使用ExcelJS - 数据修复版，兼容低版本Babel）
 */
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import ExcelStyleHelper from './excel/ExcelStyleHelper';
import ExcelTableHelper from './excel/ExcelTableHelper';

export default {
  /**
   * 将表格导出为Excel文件
   * @param {HTMLElement} editorElement - 编辑器元素
   * @param {String} fileName - 导出的文件名(可选)
   * @returns {Promise<Boolean>} - 是否成功导出
   */
  async exportTableToExcel(editorElement, fileName) {
    // 设置默认文件名
    if (!fileName) {
      const today = new Date();
      const year = today.getFullYear();
      const month = String(today.getMonth() + 1).padStart(2, '0');
      const day = String(today.getDate()).padStart(2, '0');
      fileName = `表格导出_${year}-${month}-${day}`;
    }
    
    if (!editorElement) {
      console.error('编辑器元素不存在');
      return false;
    }
    
    try {
      // 设置查找范围以同时支持v-if和v-show
      let elementsToCheck = null;
      if (editorElement.querySelectorAll) {
        elementsToCheck = editorElement;
      } else {
        // 如果传入的不是DOM元素，则在整个document中查找
        elementsToCheck = document.body;
      }
      
      // 获取所有表格，包括隐藏的（v-show会留在DOM中但display:none）
      const tables = elementsToCheck.querySelectorAll('table');
      
      if (!tables || tables.length === 0) {
        console.error('未找到可导出的表格');
        return false;
      }

      // 创建工作簿
      const workbook = new ExcelJS.Workbook();
      
      // 处理每一个表格
      let processedCount = 0;
      
      for (let i = 0; i < tables.length; i++) {
        const table = tables[i];
        
        // 跳过嵌套表格（父表格中的子表格）
        if (ExcelTableHelper.isNestedTable(table)) {
          continue;
        }
        
        // 检查表格是否隐藏 (v-show)
        const computedStyle = window.getComputedStyle(table);
        if (computedStyle.display === 'none') {
          // 临时显示表格以便处理
          table.style.display = 'table';
          table.dataset.tempDisplay = 'true';
        }
        
        // 添加工作表
        const worksheet = workbook.addWorksheet(`表格${processedCount + 1}`);
        
        // 处理表格数据
        this.processTable(table, worksheet);
        
        // 恢复临时显示的表格
        if (table.dataset.tempDisplay === 'true') {
          table.style.display = 'none';
          delete table.dataset.tempDisplay;
        }
        
        processedCount++;
      }
      
      if (processedCount === 0) {
        console.error('没有找到可处理的顶级表格');
        return false;
      }
      
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
    if (!table || !worksheet) {
      return;
    }
    
    const rows = table.querySelectorAll('tr');
    const skipCells = {};
    
    Array.from(rows).forEach(function(row, rowIndex) {
      const cells = row.querySelectorAll('th, td');
      let colIndex = 1; // ExcelJS中列索引从1开始
      
      Array.from(cells).forEach(function(cell) {
        // 检查是否需要跳过当前单元格
        while (skipCells[`${rowIndex + 1},${colIndex}`]) {
          colIndex++;
        }
        
        // 获取单元格内容
        const text = ExcelTableHelper.getCellContent(cell);
        
        // 设置单元格值
        const excelCell = worksheet.getCell(rowIndex + 1, colIndex);
        excelCell.value = text;
        
        // 应用样式
        ExcelStyleHelper.applyCellStyle(cell, excelCell);
        
        // 处理合并单元格
        ExcelTableHelper.handleMergedCells(cell, worksheet, rowIndex, colIndex, skipCells);
        
        colIndex += parseInt(cell.getAttribute('colspan')) || 1;
      });
    });
    
    // 设置列宽和行高
    this.optimizeTableDimensions(worksheet);
  },
  
  /**
   * 优化表格尺寸
   * @param {Object} worksheet - 工作表对象
   */
  optimizeTableDimensions(worksheet) {
    if (!worksheet) {
      return;
    }
    
    this.setColumnWidths(worksheet);
    this.setRowHeights(worksheet);
  },
  
  /**
   * 设置列宽
   * @param {Object} worksheet - 工作表对象
   */
  setColumnWidths(worksheet) {
    if (!worksheet || !worksheet.columns) {
      return;
    }
    
    // 获取所有列
    const columns = worksheet.columns;
    
    for (let i = 0; i < columns.length; i++) {
      const column = columns[i];
      let maxWidth = 0;
      let maxLength = 8; // 最小列宽
      
      // 遍历该列的所有单元格，找出最长的内容
      column.eachCell({ includeEmpty: true }, function(cell) {
        if (!cell.value) return;
        
        const cellValue = cell.value.toString();
        let cellWidth = ExcelStyleHelper.calculateContentWidth(cellValue);
        maxWidth = Math.max(maxWidth, cellWidth);
        
        let length = cellValue.length;
        // 为中文字符增加额外宽度
        const chineseChars = cellValue.match(/[\u4e00-\u9fa5]/g);
        if (chineseChars) {
          length += chineseChars.length * 0.5;
        }
        
        maxLength = Math.max(maxLength, length);
      });
      
      // 设置列宽，更智能地计算宽度
      if (maxWidth > 0) {
        // 设置列宽，更好地自适应文字
        const calculatedWidth = Math.min(maxWidth * 0.14 + 2, 40); 
        column.width = calculatedWidth;
      } else {
        // 基于字符长度的备选方案
        column.width = Math.min(maxLength + 2, 40); 
      }
    }
  },
  
  /**
   * 设置行高
   * @param {Object} worksheet - 工作表对象
   */
  setRowHeights(worksheet) {
    if (!worksheet) {
      return;
    }
    
    // 遍历所有行
    worksheet.eachRow({ includeEmpty: false }, function(row) {
      let maxHeight = 20; // 默认行高
      
      // 检查内容高度
      row.eachCell({ includeEmpty: false }, function(cell) {
        if (!cell.value) return;
        
        // 如果内容包含换行符，增加行高
        const value = cell.value.toString();
        const lines = value.split('\n').length;
        const cellHeight = Math.max(lines * 18, 20); // 每行至少18px，最少20px
        
        maxHeight = Math.max(maxHeight, cellHeight);
      });
      
      row.height = maxHeight;
    });
  }
};
