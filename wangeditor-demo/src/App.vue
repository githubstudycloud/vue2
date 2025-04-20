<template>
  <div id="app">
    <h1>WangEditor 表格示例</h1>
    <div class="editor-container">
      <rich-editor ref="editor" v-model="editorContent" />
    </div>
    <div class="button-container">
      <button @click="exportToExcel">导出为Excel</button>
      <button @click="exportToExcelPro" class="pro-button">导出为Excel (增强版)</button>
    </div>
  </div>
</template>

<script>
import RichEditor from './components/RichEditor.vue';
import ExcelExporter from './utils/ExcelExporter';
import ExcelExporterPro from './utils/ExcelExporterPro';

export default {
  name: 'App',
  components: {
    RichEditor
  },
  data() {
    return {
      editorContent: ''
    }
  },
  methods: {
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

button {
  background-color: #4CAF50;
  border: none;
  color: white;
  padding: 10px 20px;
  text-align: center;
  text-decoration: none;
  display: inline-block;
  font-size: 16px;
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
</style>