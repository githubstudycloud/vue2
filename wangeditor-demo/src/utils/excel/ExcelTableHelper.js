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
    
    // 检查父元素是否包含表格
    let parent = table.parentElement;
    while (parent) {
      if (parent.tagName && parent.tagName.toUpperCase() === 'TABLE') {
        return true;
      }
      
      // 检查TD或TH单元格 - 这通常表示嵌套表格
      if (parent.tagName && (parent.tagName.toUpperCase() === 'TD' || parent.tagName.toUpperCase() === 'TH')) {
        let cellParent = parent.parentElement;
        while (cellParent) {
          if (cellParent.tagName && cellParent.tagName.toUpperCase() === 'TABLE') {
            return true;
          }
          cellParent = cellParent.parentElement;
        }
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
    
    try {
      // 检查单元格是否为空
      if (!cell.textContent || cell.textContent.trim() === '') {
        return ' '; // 返回空格而非空字符串，确保表格边框正确显示
      }
      
      // 处理包含嵌套表格的单元格
      const nestedTable = cell.querySelector('table');
      if (nestedTable) {
        // 将嵌套表格转换为文本内容
        return this.convertNestedTableToText(nestedTable);
      }
      
      // 检查是否包含复杂内容（图片、列表等）
      if (cell.querySelector('img, ul, ol')) {
        // 处理包含特殊元素的单元格
        return this.extractFormattedContent(cell);
      }
      
      // 获取单元格文本内容，保留基本格式
      return cell.textContent.trim() || ' ';
    } catch (error) {
      console.error('获取单元格内容失败:', error);
      return cell.textContent ? cell.textContent.trim() : ' ';
    }
  },
  
  /**
   * 提取单元格格式化内容
   * @param {HTMLElement} cell - 单元格元素
   * @returns {String} - 格式化内容
   */
  extractFormattedContent(cell) {
    if (!cell) {
      return '';
    }
    
    let content = '';
    
    // 处理图片
    const images = cell.querySelectorAll('img');
    if (images && images.length > 0) {
      Array.from(images).forEach(function(img) {
        content += '[Image]';
      });
    }
    
    // 处理列表
    const lists = cell.querySelectorAll('ul, ol');
    if (lists && lists.length > 0) {
      Array.from(lists).forEach(function(list) {
        const listItems = list.querySelectorAll('li');
        Array.from(listItems).forEach(function(item, index) {
          if (list.tagName.toLowerCase() === 'ol') {
            content += (index + 1) + '. ' + item.textContent.trim() + '\n';
          } else {
            content += '- ' + item.textContent.trim() + '\n';
          }
        });
      });
    }
    
    // 如果没有上述特殊内容，则直接返回文本
    if (!content) {
      content = cell.textContent.trim();
    }
    
    return content || ' ';
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
    this.applyMergeCells(worksheet, rowIndex, colIndex, rowspan, colspan);
    this.markSkippedCells(skipCells, rowIndex, colIndex, rowspan, colspan);
  }
},

/**
 * 应用合并单元格
 * @param {Object} worksheet - 工作表对象
 * @param {Number} rowIndex - 行索引
 * @param {Number} colIndex - 列索引
 * @param {Number} rowspan - 行合并数量
 * @param {Number} colspan - 列合并数量
 */
applyMergeCells(worksheet, rowIndex, colIndex, rowspan, colspan) {
  worksheet.mergeCells(
    rowIndex + 1,
    colIndex,
    rowIndex + rowspan,
    colIndex + colspan - 1
  );
},

/**
 * 标记被合并的单元格
 * @param {Object} skipCells - 跳过的单元格记录
 * @param {Number} rowIndex - 行索引
 * @param {Number} colIndex - 列索引
 * @param {Number} rowspan - 行合并数量
 * @param {Number} colspan - 列合并数量
 */
markSkippedCells(skipCells, rowIndex, colIndex, rowspan, colspan) {
  for (let i = 0; i < rowspan; i++) {
    for (let j = 0; j < colspan; j++) {
      if (i !== 0 || j !== 0) {
        skipCells[`${rowIndex + 1 + i},${colIndex + j}`] = true;
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
      console.warn('预处理表格: 传入的表格元素为空');
      return;
    }
    
    console.log('开始预处理表格...');
    
    try {
      // 验证表格基本结构
      if (!table.tagName || table.tagName.toUpperCase() !== 'TABLE') {
        console.warn('预处理表格: 元素不是表格');
        return;
      }
      
      console.log('- 检查表格结构有效性');
      
      // 确保表格有tbody
      if (!table.querySelector('tbody')) {
        console.log('- 表格缺少tbody元素，创建它');
        const tbody = document.createElement('tbody');
        // 将所有没有包裹在tbody中的tr移动到新tbody中
        const rows = table.querySelectorAll('tr');
        let needsReparent = false;
        
        Array.from(rows).forEach(function(row) {
          if (row.parentElement.tagName.toUpperCase() !== 'TBODY') {
            needsReparent = true;
            tbody.appendChild(row);
          }
        });
        
        if (needsReparent) {
          table.appendChild(tbody);
        }
      }
      
      console.log('- 处理空单元格');
      // 处理空单元格
      const cells = table.querySelectorAll('td, th');
      Array.from(cells).forEach(function(cell) {
        if ((!cell.textContent || cell.textContent.trim() === '') && !cell.querySelector('table, img, ul, ol')) {
          // 为空单元格添加空格，确保Excel中显示边框
          cell.innerHTML = '&nbsp;';
        }
      });
      
      console.log('- 检查和修复行列结构');
      // 检查是否有不完整的行/列（修复一些情况下表格结构不完整的问题）
      const rows = table.querySelectorAll('tr');
      if (rows.length > 0) {
        // 查找最大列数
        let maxCols = 0;
        let totalCellsCount = 0;
        
        Array.from(rows).forEach(function(row) {
          const rowCells = row.querySelectorAll('td, th');
          let effectiveColCount = 0;
          
          // 计算有效列数，考虑colspan
          Array.from(rowCells).forEach(function(cell) {
            const colspan = parseInt(cell.getAttribute('colspan')) || 1;
            effectiveColCount += colspan;
          });
          
          maxCols = Math.max(maxCols, effectiveColCount);
          totalCellsCount += rowCells.length;
        });
        
        console.log(`- 表格有 ${rows.length} 行, ${totalCellsCount} 个单元格, 最大列数 ${maxCols}`);
        
        // 为不完整的行添加缺失的单元格
        Array.from(rows).forEach(function(row, rowIndex) {
          const rowCells = row.querySelectorAll('td, th');
          let effectiveColCount = 0;
          
          // 计算当前行的有效列数
          Array.from(rowCells).forEach(function(cell) {
            const colspan = parseInt(cell.getAttribute('colspan')) || 1;
            effectiveColCount += colspan;
          });
          
          // 如果列数不足，添加缺失的单元格
          if (effectiveColCount < maxCols) {
            console.log(`- 第 ${rowIndex+1} 行列数不足，添加 ${maxCols - effectiveColCount} 个单元格`);
            for (let i = effectiveColCount; i < maxCols; i++) {
              const newCell = document.createElement('td');
              newCell.innerHTML = '&nbsp;';
              row.appendChild(newCell);
            }
          }
        });
      } else {
        console.warn('- 表格没有行元素');
      }
      
      // 处理潜在的嵌套表格问题
      const nestedTables = table.querySelectorAll('table');
      if (nestedTables && nestedTables.length > 0) {
        console.log(`- 发现 ${nestedTables.length} 个嵌套表格`);
      }
      
      console.log('表格预处理完成');
    } catch (error) {
      console.error('预处理表格时出错:', error);
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

  // 解析宽度值为数字
  const parseWidth = (widthStr) => {
    if (!widthStr) return 10; // 默认宽度

    if (widthStr.endsWith('px')) {
      return parseFloat(widthStr) || 10;
    } else if (widthStr.endsWith('%')) {
      // 百分比宽度转换为相对值，基准为60
      return (parseFloat(widthStr) * 0.6) || 10;
    }
    return parseFloat(widthStr) || 10;
  };

  // 处理元素集合中的宽度
  const processElementWidths = (elements) => {
    Array.from(elements).forEach(element => {
      const width = element.style.width || '';
      colWidths.push(parseWidth(width));
    });
  };

  // 优先使用col元素的宽度
  if (cols.length > 0) {
    processElementWidths(cols);
  } else {
    // 如果没有col元素，使用第一行单元格的宽度
    const firstRow = table.querySelector('tr');
    if (firstRow) {
      const cells = firstRow.querySelectorAll('th, td');
      processElementWidths(cells);
    }
  }

  return colWidths;
}
};
