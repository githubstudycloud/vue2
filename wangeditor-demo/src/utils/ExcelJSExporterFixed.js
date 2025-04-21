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
    // 存储合并单元格的信息，用于列宽计算
    const mergedCellsInfo = {};
    
    // 处理所有行
    this.processTableRows(rows, worksheet, skipCells, mergedCellsInfo);
    
    // 设置列宽和行高，传入合并单元格信息
    this.optimizeTableDimensions(worksheet, mergedCellsInfo);
  },
  
  /**
   * 处理表格的所有行
   * @param {NodeList} rows - 行元素集合
   * @param {Object} worksheet - 工作表对象
   * @param {Object} skipCells - 跳过的单元格记录
   * @param {Object} mergedCellsInfo - 合并单元格信息
   */
  processTableRows(rows, worksheet, skipCells, mergedCellsInfo) {
    Array.from(rows).forEach(function(row, rowIndex) {
      const cells = row.querySelectorAll('th, td');
      let colIndex = 1; // ExcelJS中列索引从1开始
      
      // 处理行中的所有单元格
      this.processRowCells(cells, worksheet, rowIndex, colIndex, skipCells, mergedCellsInfo);
    }, this);
  },
  
  /**
   * 处理行中的所有单元格
   * @param {NodeList} cells - 单元格元素集合
   * @param {Object} worksheet - 工作表对象
   * @param {Number} rowIndex - 行索引
   * @param {Number} startColIndex - 起始列索引
   * @param {Object} skipCells - 跳过的单元格记录
   * @param {Object} mergedCellsInfo - 合并单元格信息
   */
  processRowCells(cells, worksheet, rowIndex, startColIndex, skipCells, mergedCellsInfo) {
    let colIndex = startColIndex;
    
    Array.from(cells).forEach(function(cell) {
      // 检查是否需要跳过当前单元格
      colIndex = this.findNextAvailableColumn(rowIndex, colIndex, skipCells);
      
      // 处理单个单元格
      this.processCell(cell, worksheet, rowIndex, colIndex, skipCells, mergedCellsInfo);
      
      // 更新列索引
      colIndex += parseInt(cell.getAttribute('colspan')) || 1;
    }, this);
  },
  
  /**
   * 查找下一个可用列
   * @param {Number} rowIndex - 行索引
   * @param {Number} startColIndex - 起始列索引
   * @param {Object} skipCells - 跳过的单元格记录
   * @returns {Number} - 下一个可用列索引
   */
  findNextAvailableColumn(rowIndex, startColIndex, skipCells) {
    let colIndex = startColIndex;
    while (skipCells[`${rowIndex + 1},${colIndex}`]) {
      colIndex++;
    }
    return colIndex;
  },
  
  /**
   * 处理单个单元格
   * @param {HTMLElement} cell - 单元格元素
   * @param {Object} worksheet - 工作表对象
   * @param {Number} rowIndex - 行索引
   * @param {Number} colIndex - 列索引
   * @param {Object} skipCells - 跳过的单元格记录
   * @param {Object} mergedCellsInfo - 合并单元格信息
   */
  processCell(cell, worksheet, rowIndex, colIndex, skipCells, mergedCellsInfo) {
    // 获取单元格内容
    const text = ExcelTableHelper.getCellContent(cell);
    
    // 设置单元格值
    const excelCell = worksheet.getCell(rowIndex + 1, colIndex);
    excelCell.value = text;
    
    // 应用样式
    ExcelStyleHelper.applyCellStyle(cell, excelCell);
    
    // 记录合并单元格信息
    const colspan = parseInt(cell.getAttribute('colspan')) || 1;
    if (colspan > 1) {
      this.recordMergedCellInfo(mergedCellsInfo, colIndex, colspan, text);
    }
    
    // 处理合并单元格
    ExcelTableHelper.handleMergedCells(cell, worksheet, rowIndex, colIndex, skipCells);
  },
  
  /**
   * 记录合并单元格信息
   * @param {Object} mergedCellsInfo - 合并单元格信息集合
   * @param {Number} colIndex - 列索引
   * @param {Number} colspan - 合并列数
   * @param {String} text - 单元格内容
   */
  recordMergedCellInfo(mergedCellsInfo, colIndex, colspan, text) {
    // 计算单元格内容宽度
    const width = ExcelStyleHelper.calculateContentWidth(text);
    
    // 为每一个被合并的列记录信息
    for (let i = colIndex; i < colIndex + colspan; i++) {
      if (!mergedCellsInfo[i]) {
        mergedCellsInfo[i] = [];
      }
      
      // 添加合并单元格信息
      mergedCellsInfo[i].push({
        content: text,
        width: width,
        colspan: colspan
      });
    }
  },
  
  /**
   * 优化表格尺寸
   * @param {Object} worksheet - 工作表对象
   * @param {Object} mergedCellsInfo - 合并单元格信息
   */
  optimizeTableDimensions(worksheet, mergedCellsInfo) {
    if (!worksheet) {
      return;
    }
    
    this.setColumnWidths(worksheet, mergedCellsInfo);
    this.setRowHeights(worksheet);
  },
  
  /**
   * 设置列宽
   * @param {Object} worksheet - 工作表对象
   * @param {Object} mergedCellsInfo - 合并单元格信息
   */
  setColumnWidths(worksheet, mergedCellsInfo) {
    if (!worksheet || !worksheet.columns) {
      return;
    }
    
    // 获取所有列
    const columns = worksheet.columns;
    
    for (let i = 0; i < columns.length; i++) {
      const column = columns[i];
      const colIndex = i + 1; // ExcelJS列索引从1开始
      
      // 收集列宽度数据
      const widthData = this.collectColumnWidthData(column, colIndex, mergedCellsInfo);
      
      // 计算并应用最终列宽
      this.applyColumnWidth(column, widthData.maxWidth, widthData.maxLength);
    }
  },
  
  /**
   * 收集列宽度数据
   * @param {Object} column - 列对象
   * @param {Number} colIndex - 列索引
   * @param {Object} mergedCellsInfo - 合并单元格信息
   * @returns {Object} - 宽度数据
   */
  collectColumnWidthData(column, colIndex, mergedCellsInfo) {
    let maxWidth = 0;
    let maxLength = 8; // 最小列宽
    
    // 处理普通单元格
    this.processRegularCells(column, function(cellValue) {
      const cellWidth = ExcelStyleHelper.calculateContentWidth(cellValue);
      maxWidth = Math.max(maxWidth, cellWidth);
      
      let length = cellValue.length;
      // 为中文字符增加额外宽度
      const chineseChars = cellValue.match(/[\u4e00-\u9fa5]/g);
      if (chineseChars) {
        length += chineseChars.length * 0.5;
      }
      
      maxLength = Math.max(maxLength, length);
    });
    
    // 处理合并单元格
    this.processMergedCells(colIndex, mergedCellsInfo, function(adjustedWidth) {
      maxWidth = Math.max(maxWidth, adjustedWidth);
    });
    
    return { maxWidth, maxLength };
  },
  
  /**
   * 处理普通单元格
   * @param {Object} column - 列对象
   * @param {Function} callback - 处理回调
   */
  processRegularCells(column, callback) {
    column.eachCell({ includeEmpty: true }, function(cell) {
      if (!cell || !cell.value) return;
      
      const cellValue = cell.value.toString();
      callback(cellValue);
    });
  },
  
  /**
   * 处理合并单元格
   * @param {Number} colIndex - 列索引
   * @param {Object} mergedCellsInfo - 合并单元格信息
   * @param {Function} callback - 处理回调
   */
  processMergedCells(colIndex, mergedCellsInfo, callback) {
    if (!mergedCellsInfo || !mergedCellsInfo[colIndex]) {
      return;
    }
    
    const mergedCells = mergedCellsInfo[colIndex];
    for (let j = 0; j < mergedCells.length; j++) {
      const mergedInfo = mergedCells[j];
      // 计算合并单元格的平均宽度
      if (mergedInfo.colspan > 1) {
        // 如果是合并单元格，宽度需要除以合并的列数
        const adjustedWidth = mergedInfo.width / mergedInfo.colspan;
        callback(adjustedWidth);
      }
    }
  },
  
  /**
   * 应用最终列宽
   * @param {Object} column - 列对象
   * @param {Number} maxWidth - 最大内容宽度
   * @param {Number} maxLength - 最大内容长度
   */
  applyColumnWidth(column, maxWidth, maxLength) {
    if (maxWidth > 0) {
      // 设置列宽，更好地自适应文字
      // 有一个转换因子(0.14)将像素转换为Excel单位，并加上一点额外空间(+2)
      // 同时限制最大宽度为40，防止过宽
      const calculatedWidth = Math.min(maxWidth * 0.14 + 2, 40);
      column.width = calculatedWidth;
    } else {
      // 基于字符长度的备选方案
      column.width = Math.min(maxLength + 2, 40);
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
    
    // 使用数据收集器先收集所有行的高度
    const rowHeights = this.collectRowHeights(worksheet);
    
    // 然后应用高度
    this.applyRowHeights(worksheet, rowHeights);
  },
  
  /**
   * 收集行高数据
   * @param {Object} worksheet - 工作表对象
   * @returns {Object} - 行高数据映射
   */
  collectRowHeights(worksheet) {
    const rowHeights = {};
    
    worksheet.eachRow({ includeEmpty: false }, function(row) {
      // 收集当前行的最大高度
      const maxHeight = this.calculateRowMaxHeight(row);
      rowHeights[row.number] = maxHeight;
    }, this);
    
    return rowHeights;
  },
  
  /**
   * 计算行的最大高度
   * @param {Object} row - 行对象
   * @returns {Number} - 行高
   */
  calculateRowMaxHeight(row) {
    let maxHeight = 20; // 默认行高
    
    row.eachCell({ includeEmpty: false }, function(cell) {
      if (!cell || !cell.value) return;
      
      // 获取单元格的内容高度
      const cellHeight = this.calculateCellHeight(cell);
      maxHeight = Math.max(maxHeight, cellHeight);
    }, this);
    
    return maxHeight;
  },
  
  /**
   * 计算单元格高度
   * @param {Object} cell - 单元格对象
   * @returns {Number} - 单元格高度
   */
  calculateCellHeight(cell) {
    // 获取单元格内容
    const value = cell.value ? cell.value.toString() : '';
    
    // 计算行数
    const lines = value.split('\n').length;
    
    // 每行至少18px，最少20px
    return Math.max(lines * 18, 20);
  },
  
  /**
   * 应用行高
   * @param {Object} worksheet - 工作表对象
   * @param {Object} rowHeights - 行高数据
   */
  applyRowHeights(worksheet, rowHeights) {
    // 遍历所有行，应用高度
    worksheet.eachRow({ includeEmpty: false }, function(row) {
      const rowNumber = row.number;
      if (rowHeights[rowNumber]) {
        row.height = rowHeights[rowNumber];
      }
    });
  }
};
