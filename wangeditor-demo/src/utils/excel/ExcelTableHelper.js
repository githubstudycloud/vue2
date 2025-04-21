/**
 * Excel表格工具类 - 负责处理表格内容和结构
 */

export default {
  /**
   * 检查是否是嵌套表格
   * @param {HTMLTableElement} table - 表格元素
   * @returns {Boolean} - 是否是嵌套表格
   */
  isNestedTable(table) {
    if (!table) {
      return false;
    }
    
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
   * 获取单元格内容
   * @param {HTMLElement} cell - 单元格元素
   * @returns {String} - 单元格内容
   */
  getCellContent(cell) {
    if (!cell) {
      return '';
    }
    
    // 处理包含嵌套表格的单元格
    const nestedTable = cell.querySelector('table');
    if (nestedTable) {
      // 将嵌套表格转换为文本内容
      return this.convertNestedTableToText(nestedTable);
    } else {
      // 获取单元格文本内容
      return cell.textContent.trim();
    }
  },
  
  /**
   * 将嵌套表格转换为文本内容
   * @param {HTMLTableElement} table - 嵌套表格元素
   * @returns {String} - 文本内容
   */
  convertNestedTableToText(table) {
    if (!table) {
      return '';
    }
    
    let text = '';
    const rows = table.querySelectorAll('tr');
    
    Array.from(rows).forEach(function(row, rowIndex) {
      const cells = row.querySelectorAll('th, td');
      let rowText = '';
      
      Array.from(cells).forEach(function(cell, cellIndex) {
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
   * 处理合并单元格
   * @param {HTMLElement} cell - 单元格元素
   * @param {Object} worksheet - 工作表对象
   * @param {Number} rowIndex - 行索引
   * @param {Number} colIndex - 列索引
   * @param {Object} skipCells - 跳过的单元格记录
   */
  handleMergedCells(cell, worksheet, rowIndex, colIndex, skipCells) {
    if (!cell || !worksheet) {
      return;
    }
    
    const colspan = parseInt(cell.getAttribute('colspan')) || 1;
    const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
    
    if (colspan > 1 || rowspan > 1) {
      // 应用合并单元格
      worksheet.mergeCells(
        rowIndex + 1,
        colIndex,
        rowIndex + rowspan,
        colIndex + colspan - 1
      );
      
      // 标记被合并的单元格，以便跳过处理
      for (let i = 0; i < rowspan; i++) {
        for (let j = 0; j < colspan; j++) {
          if (i !== 0 || j !== 0) {
            skipCells[`${rowIndex + 1 + i},${colIndex + j}`] = true;
          }
        }
      }
    }
  },
  
  /**
   * 查找v-if隐藏的表格
   * 注: 这个方法用于在DOM中查找通过v-if条件渲染的表格
   * 当表格使用v-if时，表格可能不在当前DOM树中，需要通过特定方式查找
   * @param {HTMLElement} container - 容器元素
   * @returns {Array} - 找到的表格数组
   */
  findHiddenTables(container) {
    if (!container) {
      return [];
    }
    
    // 获取所有具有特定属性的元素，这些元素可能是v-if控制的容器
    const vIfContainers = container.querySelectorAll('[v-if]');
    const tables = [];
    
    Array.from(vIfContainers).forEach(function(elem) {
      // 检查元素是否为表格或包含表格
      if (elem.tagName === 'TABLE') {
        tables.push(elem);
      } else {
        const nestedTables = elem.querySelectorAll('table');
        if (nestedTables && nestedTables.length > 0) {
          Array.from(nestedTables).forEach(function(table) {
            tables.push(table);
          });
        }
      }
    });
    
    return tables;
  },
  
  /**
   * 查找并临时显示v-show隐藏的表格
   * @param {HTMLElement} container - 容器元素
   * @returns {Array} - 临时显示的表格元素数组
   */
  showHiddenVShowTables(container) {
    if (!container) {
      return [];
    }
    
    // 查找所有表格
    const tables = container.querySelectorAll('table');
    const tempShownTables = [];
    
    // 检查每个表格的显示状态
    Array.from(tables).forEach(function(table) {
      const computedStyle = window.getComputedStyle(table);
      if (computedStyle.display === 'none') {
        // 临时显示表格
        table.style.display = 'table';
        table.dataset.tempDisplay = 'true';
        tempShownTables.push(table);
      }
    });
    
    return tempShownTables;
  },
  
  /**
   * 恢复临时显示的表格
   * @param {Array} tables - 临时显示的表格数组
   */
  restoreHiddenTables(tables) {
    if (!tables || !tables.length) {
      return;
    }
    
    tables.forEach(function(table) {
      if (table.dataset.tempDisplay === 'true') {
        table.style.display = 'none';
        delete table.dataset.tempDisplay;
      }
    });
  },
  
  /**
   * 预处理表格，清理一些可能导致导出问题的结构
   * @param {HTMLTableElement} table - 表格元素
   */
  preprocessTable(table) {
    if (!table) {
      return;
    }
    
    // 处理空单元格
    const cells = table.querySelectorAll('td, th');
    Array.from(cells).forEach(function(cell) {
      if (!cell.textContent.trim() && !cell.querySelector('table')) {
        // 为空单元格添加空格，确保Excel中显示边框
        cell.innerHTML = '&nbsp;';
      }
    });
    
    // 检查是否有不完整的行/列（修复一些情况下表格结构不完整的问题）
    const rows = table.querySelectorAll('tr');
    if (rows.length > 0) {
      // 查找最大列数
      let maxCols = 0;
      Array.from(rows).forEach(function(row) {
        const cols = row.querySelectorAll('td, th').length;
        maxCols = Math.max(maxCols, cols);
      });
      
      // 为不完整的行添加缺失的单元格
      Array.from(rows).forEach(function(row) {
        const cols = row.querySelectorAll('td, th').length;
        if (cols < maxCols) {
          for (let i = cols; i < maxCols; i++) {
            const newCell = document.createElement('td');
            newCell.innerHTML = '&nbsp;';
            row.appendChild(newCell);
          }
        }
      });
    }
  },
  
  /**
   * 处理表格列宽度
   * @param {HTMLTableElement} table - 表格元素
   * @returns {Array} - 列宽度数组
   */
  getTableColumnWidths(table) {
    if (!table) {
      return [];
    }
    
    // 获取表格的列宽度信息
    const colWidths = [];
    const cols = table.querySelectorAll('col');
    
    // 如果有col元素，使用这些信息
    if (cols.length > 0) {
      Array.from(cols).forEach(function(col) {
        let width = col.style.width || '';
        // 将宽度值转换为数字
        if (width) {
          if (width.endsWith('px')) {
            width = parseFloat(width);
          } else if (width.endsWith('%')) {
            // 百分比宽度转换为相对值，基准为60
            width = parseFloat(width) * 0.6;
          } else {
            width = parseFloat(width);
          }
          colWidths.push(width || 10); // 默认宽度为10
        } else {
          colWidths.push(10); // 默认宽度
        }
      });
    } else {
      // 如果没有col元素，检查第一行的单元格宽度
      const firstRow = table.querySelector('tr');
      if (firstRow) {
        const cells = firstRow.querySelectorAll('th, td');
        Array.from(cells).forEach(function(cell) {
          let width = cell.style.width || '';
          if (width) {
            if (width.endsWith('px')) {
              width = parseFloat(width);
            } else if (width.endsWith('%')) {
              width = parseFloat(width) * 0.6;
            } else {
              width = parseFloat(width);
            }
            colWidths.push(width || 10);
          } else {
            colWidths.push(10);
          }
        });
      }
    }
    
    return colWidths;
  }
};
