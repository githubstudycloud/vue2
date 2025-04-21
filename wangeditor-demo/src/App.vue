<template>
  <div id="app">
    <h1>WangEditor 表格示例</h1>
    <div class="editor-container">
      <rich-editor ref="editor" v-model="editorContent" />
    </div>
    <div class="button-container">
      <div class="example-buttons">
        <button @click="switchExample('default')" :class="{ active: currentExample === 'default' }">默认示例</button>
        <button @click="switchExample('projectProgress')" :class="{ active: currentExample === 'projectProgress' }">项目进度表</button>
        <button @click="switchExample('hrDepartment')" :class="{ active: currentExample === 'hrDepartment' }">人力资源表</button>
        <button @click="switchExample('financialReport')" :class="{ active: currentExample === 'financialReport' }">财务报表</button>
        <button @click="switchExample('productComparison')" :class="{ active: currentExample === 'productComparison' }">产品对比表</button>
      </div>
      <div class="export-buttons">
        <button @click="exportToExcel">导出为Excel</button>
        <button @click="exportToExcelPro" class="pro-button">导出为Excel (增强版)</button>
        <button @click="exportWithExcelJS" class="exceljs-button">导出为Excel (ExcelJS版)</button>
        <button @click="exportWithExcelJSFixed" class="fixed-button">导出为Excel (修复版)</button>
      </div>
    </div>
  </div>
</template>

<script>
import RichEditor from './components/RichEditor.vue';
import ExcelExporter from './utils/ExcelExporter';
import ExcelExporterPro from './utils/ExcelExporterPro';
import ExcelJSExporter from './utils/ExcelJSExporter';
import ExcelJSExporterFixed from './utils/ExcelJSExporterFixed';
import * as AdditionalTableExamples from './utils/AdditionalTableExamples';

export default {
  name: 'App',
  components: {
    RichEditor
  },
  data() {
    return {
      editorContent: '',
      currentExample: 'default'
    }
  },
  mounted() {
    // 页面加载时，默认显示默认示例
    this.currentExample = 'default';
  },
  methods: {
    switchExample(exampleType) {
      if (this.currentExample === exampleType) {
        return; // 如果是同一个示例，不需要切换
      }
      
      let content = '';
      
      switch (exampleType) {
        case 'default':
          // 使用默认示例
          content = '';
          break;
        case 'projectProgress':
          content = AdditionalTableExamples.getProjectProgressTableHTML();
          break;
        case 'hrDepartment':
          content = AdditionalTableExamples.getHRDepartmentTableHTML();
          break;
        case 'financialReport':
          content = AdditionalTableExamples.getFinancialReportTableHTML();
          break;
        case 'productComparison':
          content = AdditionalTableExamples.getProductComparisonTableHTML();
          break;
      }
      
      this.currentExample = exampleType;
      
      // 确保编辑器已经初始化
      this.$nextTick(() => {
        if (this.$refs.editor) {
          this.$refs.editor.setContent(content);
        }
      });
    },
    
    exportToExcel() {
      if (!this.$refs.editor) {
        alert('编辑器不可用');
        return;
      }
      
      const editorElement = this.$refs.editor.getEditorElement();
      const success = ExcelExporter.exportTableToExcel(editorElement);
      
      if (!success) {
        alert('导出失败，请确保编辑器中有表格内容');
      }
    },
    
    exportToExcelPro() {
      if (!this.$refs.editor) {
        alert('编辑器不可用');
        return;
      }
      
      const editorElement = this.$refs.editor.getEditorElement();
      const success = ExcelExporterPro.exportTableToExcel(editorElement);
      
      if (!success) {
        alert('导出失败，请确保编辑器中有表格内容');
      }
    },
    
    async exportWithExcelJS() {
      if (!this.$refs.editor) {
        alert('编辑器不可用');
        return;
      }
      
      const editorElement = this.$refs.editor.getEditorElement();
      const success = await ExcelJSExporter.exportTableToExcel(editorElement);
      
      if (!success) {
        alert('导出失败，请确保编辑器中有表格内容');
      }
    },
    
    async exportWithExcelJSFixed() {
      if (!this.$refs.editor) {
        alert('编辑器不可用');
        return;
      }
      
      console.log('===========================');
      console.log('开始导出Excel (修复版)');
      console.log('===========================');
      
      const editorElement = this.$refs.editor.getEditorElement();
      const success = await ExcelJSExporterFixed.exportTableToExcel(editorElement);
      
      console.log('===========================');
      console.log('导出完成，结果:', success ? '成功' : '失败');
      console.log('===========================');
      
      if (!success) {
        alert('导出失败，请确保编辑器中有表格内容');
      }
    }
  }
}
</script>

<style>
#app {
  font-family: 'Microsoft YaHei', Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin: 20px auto;
  max-width: 1200px;
  padding: 0 20px;
}

.editor-container {
  margin: 20px 0;
  border: 1px solid #ccc;
  border-radius: 5px;
  overflow: hidden;
  text-align: left;
}

.button-container {
  margin: 20px 0;
}

.example-buttons, .export-buttons {
  display: flex;
  justify-content: center;
  flex-wrap: wrap;
  gap: 10px;
  margin-bottom: 15px;
}

.example-buttons button.active {
  background-color: #2196F3;
  color: white;
}

.example-buttons button.active:hover {
  background-color: #1976D2;
}

button {
  background-color: #4CAF50;
  border: none;
  color: white;
  padding: 10px 20px;
  text-align: center;
  text-decoration: none;
  display: inline-block;
  font-size: 14px;
  margin: 4px 2px;
  cursor: pointer;
  border-radius: 4px;
}

button:hover {
  background-color: #45a049;
}

.pro-button {
  background-color: #2196F3;
}

.pro-button:hover {
  background-color: #1976D2;
}

.exceljs-button {
  background-color: #9C27B0;
}

.exceljs-button:hover {
  background-color: #7B1FA2;
}

.fixed-button {
  background-color: #FF5722;
}

.fixed-button:hover {
  background-color: #E64A19;
}
</style>