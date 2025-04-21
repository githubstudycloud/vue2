/**
 * 表格样式工具类 - 提供表格样式增强功能
 */

export default {
  /**
   * 为表格添加背景色，并确保背景色占满整个表格
   * @param {HTMLElement} table - 表格元素
   * @param {Object} options - 配置选项
   * @param {String} options.backgroundColor - 背景色 (CSS颜色值)
   * @param {String} options.borderColor - 边框颜色 (CSS颜色值)
   * @param {String} options.borderWidth - 边框宽度 (CSS宽度值，如'1px')
   * @param {String} options.borderStyle - 边框样式 (如'solid', 'dashed')
   * @returns {Boolean} - 是否操作成功
   */
  applyTableBackground(table, options = {}) {
    if (!table || table.tagName !== 'TABLE') {
      console.error('无效的表格元素');
      return false;
    }
    
    // 默认选项
    const defaultOptions = {
      backgroundColor: '#f5f5f5',
      borderColor: '#dddddd',
      borderWidth: '1px',
      borderStyle: 'solid'
    };
    
    // 合并选项
    const mergedOptions = Object.assign({}, defaultOptions, options);
    
    try {
      // 设置表格基本样式
      this.setTableBaseStyle(table, mergedOptions);
      
      // 确保每个单元格都有边框和背景
      this.setTableCellsStyle(table, mergedOptions);
      
      // 处理合并单元格
      this.handleMergedCells(table, mergedOptions);
      
      // 处理表格外围边框
      this.setTableOuterBorder(table, mergedOptions);
      
      return true;
    } catch (error) {
      console.error('应用表格背景时出错:', error);
      return false;
    }
  },
  
  /**
   * 设置表格基本样式
   * @param {HTMLElement} table - 表格元素
   * @param {Object} options - 样式选项
   */
  setTableBaseStyle(table, options) {
    // 设置表格基本样式
    table.style.width = '100%'; // 确保表格宽度为100%
    table.style.borderCollapse = 'collapse'; // 折叠边框
    table.style.borderSpacing = '0'; // 移除边框间距
    table.style.backgroundColor = options.backgroundColor; // 设置背景色
    table.style.border = `${options.borderWidth} ${options.borderStyle} ${options.borderColor}`; // 设置边框
    
    // 移除表格默认边距
    table.style.margin = '0';
    table.style.padding = '0';
    
    // 添加自定义数据属性，用于标记已处理的表格
    table.dataset.stylized = 'true';
  },
  
  /**
   * 设置表格单元格样式
   * @param {HTMLElement} table - 表格元素
   * @param {Object} options - 样式选项
   */
  setTableCellsStyle(table, options) {
    // 获取所有单元格
    const cells = table.querySelectorAll('th, td');
    
    // 设置每个单元格的样式
    Array.from(cells).forEach(function(cell) {
      // 设置单元格边框
      cell.style.border = `${options.borderWidth} ${options.borderStyle} ${options.borderColor}`;
      
      // 如果是表头，使用稍微深一点的背景色
      if (cell.tagName === 'TH') {
        // 计算稍微深一点的背景色
        const darkerColor = this.getDarkerColor(options.backgroundColor, 10);
        cell.style.backgroundColor = darkerColor;
      } else {
        cell.style.backgroundColor = options.backgroundColor;
      }
      
      // 确保内容有一定的内边距
      if (!cell.style.padding) {
        cell.style.padding = '8px';
      }
    }.bind(this));
  },
  
  /**
   * 处理合并单元格
   * @param {HTMLElement} table - 表格元素
   * @param {Object} options - 样式选项
   */
  handleMergedCells(table, options) {
    // 获取所有单元格
    const cells = table.querySelectorAll('th, td');
    
    // 检查每个单元格是否为合并单元格
    Array.from(cells).forEach(function(cell) {
      const colspan = parseInt(cell.getAttribute('colspan')) || 1;
      const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
      
      // 如果是合并单元格，确保背景色填充整个区域
      if (colspan > 1 || rowspan > 1) {
        // 强调合并单元格的边框
        cell.style.border = `${options.borderWidth} ${options.borderStyle} ${options.borderColor}`;
        
        // 去除内部边框影响
        if (rowspan > 1) {
          cell.style.verticalAlign = 'middle'; // 垂直居中
        }
        
        // 确保合并单元格内容有足够空间
        const paddingValue = Math.max(parseInt(cell.style.padding) || 8, 8);
        cell.style.padding = `${paddingValue}px`;
      }
    });
  },
  
  /**
   * 设置表格外围边框
   * @param {HTMLElement} table - 表格元素
   * @param {Object} options - 样式选项
   */
  setTableOuterBorder(table, options) {
    // 使用稍微粗一点的边框突出表格外围
    const outerBorderWidth = (parseInt(options.borderWidth) + 1) + 'px';
    table.style.border = `${outerBorderWidth} ${options.borderStyle} ${options.borderColor}`;
    
    // 尝试获取表格第一行的单元格
    const firstRow = table.querySelector('tr');
    if (firstRow) {
      const firstRowCells = firstRow.querySelectorAll('th, td');
      // 增强第一行单元格的上边框
      Array.from(firstRowCells).forEach(function(cell) {
        cell.style.borderTop = `${outerBorderWidth} ${options.borderStyle} ${options.borderColor}`;
      });
      
      // 尝试获取表格最后一行
      const rows = table.querySelectorAll('tr');
      if (rows.length > 0) {
        const lastRow = rows[rows.length - 1];
        const lastRowCells = lastRow.querySelectorAll('th, td');
        // 增强最后一行单元格的下边框
        Array.from(lastRowCells).forEach(function(cell) {
          cell.style.borderBottom = `${outerBorderWidth} ${options.borderStyle} ${options.borderColor}`;
        });
      }
    }
  },
  
  /**
   * 获取稍微深一点的颜色（用于表头等）
   * @param {String} color - 原始颜色（十六进制或RGB）
   * @param {Number} percent - 加深的百分比 (0-100)
   * @returns {String} - 加深后的颜色
   */
  getDarkerColor(color, percent = 10) {
    if (!color) return '#d0d0d0';
    
    try {
      // 从颜色中提取RGB值
      let r, g, b;
      
      if (color.startsWith('#')) {
        // 处理十六进制颜色
        const hex = color.substring(1);
        const bigint = parseInt(hex, 16);
        r = (bigint >> 16) & 255;
        g = (bigint >> 8) & 255;
        b = bigint & 255;
      } else if (color.startsWith('rgb')) {
        // 处理RGB颜色
        const rgbValues = color.match(/\d+/g);
        if (rgbValues && rgbValues.length >= 3) {
          r = parseInt(rgbValues[0]);
          g = parseInt(rgbValues[1]);
          b = parseInt(rgbValues[2]);
        } else {
          return '#d0d0d0';
        }
      } else {
        return '#d0d0d0';
      }
      
      // 加深颜色
      const factor = 1 - percent / 100;
      r = Math.floor(r * factor);
      g = Math.floor(g * factor);
      b = Math.floor(b * factor);
      
      // 转换回十六进制颜色
      return `#${this.toHex(r)}${this.toHex(g)}${this.toHex(b)}`;
    } catch (error) {
      console.error('计算深色时出错:', error);
      return '#d0d0d0';
    }
  },
  
  /**
   * 转换数字为十六进制
   * @param {Number} num - 数字
   * @returns {String} - 十六进制字符串
   */
  toHex(num) {
    const hex = num.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  },
  
  /**
   * 重置表格样式（移除已添加的背景和边框）
   * @param {HTMLElement} table - 表格元素
   * @returns {Boolean} - 是否操作成功
   */
  resetTableStyles(table) {
    if (!table || table.tagName !== 'TABLE') {
      console.error('无效的表格元素');
      return false;
    }
    
    try {
      // 移除表格样式
      table.style.backgroundColor = '';
      table.style.border = '';
      table.style.borderCollapse = '';
      table.style.borderSpacing = '';
      
      // 移除所有单元格样式
      const cells = table.querySelectorAll('th, td');
      Array.from(cells).forEach(function(cell) {
        cell.style.backgroundColor = '';
        cell.style.border = '';
        // 保留内边距以保持布局
        //cell.style.padding = '';
      });
      
      // 移除数据属性标记
      delete table.dataset.stylized;
      
      return true;
    } catch (error) {
      console.error('重置表格样式时出错:', error);
      return false;
    }
  },
  
  /**
   * 创建样式化的表格
   * @param {Number} rows - 行数
   * @param {Number} cols - 列数
   * @param {Object} options - 样式选项
   * @returns {HTMLElement} - 创建的表格元素
   */
  createStylizedTable(rows = 3, cols = 3, options = {}) {
    const table = document.createElement('table');
    
    // 创建表格结构
    for (let i = 0; i < rows; i++) {
      const row = document.createElement('tr');
      
      for (let j = 0; j < cols; j++) {
        // 第一行使用表头
        const cell = i === 0 ? document.createElement('th') : document.createElement('td');
        cell.textContent = i === 0 ? `表头 ${j+1}` : `单元格 ${i},${j+1}`;
        row.appendChild(cell);
      }
      
      table.appendChild(row);
    }
    
    // 应用样式
    this.applyTableBackground(table, options);
    
    return table;
  }
};
