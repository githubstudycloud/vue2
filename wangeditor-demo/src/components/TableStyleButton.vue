<template>
  <div class="table-style-button">
    <button @click="openStylePanel" class="style-btn" :class="{ active: showPanel }">
      表格样式
    </button>
    <div v-show="showPanel" class="style-panel">
      <h3>表格样式设置</h3>
      <div class="style-option">
        <label>背景颜色:</label>
        <input type="color" v-model="styleOptions.backgroundColor" />
      </div>
      <div class="style-option">
        <label>边框颜色:</label>
        <input type="color" v-model="styleOptions.borderColor" />
      </div>
      <div class="style-option">
        <label>边框宽度:</label>
        <select v-model="styleOptions.borderWidth">
          <option value="1px">细 (1px)</option>
          <option value="2px">中 (2px)</option>
          <option value="3px">粗 (3px)</option>
        </select>
      </div>
      <div class="style-option">
        <label>边框样式:</label>
        <select v-model="styleOptions.borderStyle">
          <option value="solid">实线</option>
          <option value="dashed">虚线</option>
          <option value="dotted">点线</option>
        </select>
      </div>
      <div class="style-preview" 
           :style="{
             backgroundColor: styleOptions.backgroundColor,
             border: `${styleOptions.borderWidth} ${styleOptions.borderStyle} ${styleOptions.borderColor}`
           }">
        预览效果
      </div>
      <div class="style-actions">
        <button @click="applyTableStyle" class="apply-btn">应用样式</button>
        <button @click="resetTableStyle" class="reset-btn">重置样式</button>
        <button @click="showPanel = false" class="close-btn">关闭</button>
      </div>
    </div>
  </div>
</template>

<script>
import TableStyleHelper from '../utils/TableStyleHelper';

export default {
  name: 'TableStyleButton',
  props: {
    editor: {
      type: Object,
      required: true
    }
  },
  data() {
    return {
      showPanel: false,
      styleOptions: {
        backgroundColor: '#f5f5f5',
        borderColor: '#dddddd',
        borderWidth: '1px',
        borderStyle: 'solid'
      }
    };
  },
  methods: {
    openStylePanel() {
      this.showPanel = !this.showPanel;
    },
    
    applyTableStyle() {
      if (!this.editor) {
        alert('编辑器不可用');
        return;
      }
      
      // 获取编辑器元素
      const editorElement = this.editor.getEditableContainer();
      if (!editorElement) {
        alert('无法获取编辑区域');
        return;
      }
      
      // 获取选中的表格或所有表格
      let tables = [];
      const selectedTable = this.getSelectedTable();
      
      if (selectedTable) {
        // 如果有选中的表格，只应用到选中的表格
        tables = [selectedTable];
      } else {
        // 否则应用到所有表格
        tables = editorElement.querySelectorAll('table');
      }
      
      if (tables.length === 0) {
        alert('编辑器中没有找到表格');
        return;
      }
      
      // 应用样式
      let count = 0;
      Array.from(tables).forEach(table => {
        if (TableStyleHelper.applyTableBackground(table, this.styleOptions)) {
          count++;
        }
      });
      
      alert(`已成功应用样式到 ${count} 个表格`);
      this.showPanel = false;
    },
    
    resetTableStyle() {
      if (!this.editor) {
        return;
      }
      
      const editorElement = this.editor.getEditableContainer();
      if (!editorElement) {
        return;
      }
      
      // 获取选中的表格或所有表格
      let tables = [];
      const selectedTable = this.getSelectedTable();
      
      if (selectedTable) {
        // 如果有选中的表格，只重置选中的表格
        tables = [selectedTable];
      } else {
        // 否则重置所有表格
        tables = editorElement.querySelectorAll('table');
      }
      
      if (tables.length === 0) {
        return;
      }
      
      // 重置样式
      let count = 0;
      Array.from(tables).forEach(table => {
        if (TableStyleHelper.resetTableStyles(table)) {
          count++;
        }
      });
      
      alert(`已重置 ${count} 个表格的样式`);
    },
    
    getSelectedTable() {
      try {
        // 尝试获取当前选择的表格
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) {
          return null;
        }
        
        let node = selection.getRangeAt(0).commonAncestorContainer;
        // 如果选中的是文本节点，获取其父元素
        if (node.nodeType === 3) {
          node = node.parentNode;
        }
        
        // 向上寻找表格元素
        while (node && node.nodeName !== 'TABLE') {
          node = node.parentNode;
          if (!node || node.nodeName === 'BODY') {
            return null;
          }
        }
        
        return node;
      } catch (error) {
        console.error('获取选中表格时出错:', error);
        return null;
      }
    }
  }
};
</script>

<style scoped>
.table-style-button {
  position: relative;
  display: inline-block;
}

.style-btn {
  background-color: #f0f0f0;
  border: 1px solid #ddd;
  padding: 5px 10px;
  cursor: pointer;
  border-radius: 4px;
  font-size: 14px;
}

.style-btn:hover {
  background-color: #e0e0e0;
}

.style-btn.active {
  background-color: #d0d0d0;
}

.style-panel {
  position: absolute;
  top: 100%;
  left: 0;
  width: 280px;
  background-color: white;
  border: 1px solid #ccc;
  border-radius: 4px;
  box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
  padding: 15px;
  z-index: 1000;
  margin-top: 5px;
}

.style-panel h3 {
  margin-top: 0;
  margin-bottom: 15px;
  font-size: 16px;
  color: #333;
  padding-bottom: 8px;
  border-bottom: 1px solid #eee;
}

.style-option {
  margin-bottom: 12px;
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.style-option label {
  font-size: 14px;
  color: #333;
  flex: 1;
}

.style-option input[type="color"],
.style-option select {
  width: 120px;
  padding: 3px;
  border: 1px solid #ddd;
  border-radius: 3px;
}

.style-preview {
  height: 60px;
  margin: 15px 0;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 4px;
  font-size: 14px;
  color: #333;
}

.style-actions {
  display: flex;
  justify-content: space-between;
  margin-top: 15px;
}

.style-actions button {
  padding: 5px 10px;
  border: none;
  border-radius: 3px;
  cursor: pointer;
  font-size: 14px;
}

.apply-btn {
  background-color: #1890ff;
  color: white;
}

.apply-btn:hover {
  background-color: #40a9ff;
}

.reset-btn {
  background-color: #ff4d4f;
  color: white;
}

.reset-btn:hover {
  background-color: #ff7875;
}

.close-btn {
  background-color: #f0f0f0;
  color: #333;
}

.close-btn:hover {
  background-color: #e0e0e0;
}
</style>
