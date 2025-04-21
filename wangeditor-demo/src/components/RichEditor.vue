<template>
  <div class="rich-editor">
    <div style="border-bottom: 1px solid #ccc">
      <Toolbar
        class="toolbar"
        :editor="editor"
        :defaultConfig="toolbarConfig"
        :mode="mode"
      />
    </div>
    <div class="editor-content">
      <Editor
        class="editor"
        :defaultConfig="editorConfig"
        :mode="mode"
        v-model="editorHtml"
        @onChange="handleChange"
        @onCreated="handleCreated"
      />
    </div>
  </div>
</template>

<script>
import { Editor, Toolbar } from '@wangeditor/editor-for-vue';
import '@wangeditor/editor/dist/css/style.css';
import { getAdvancedNestedTableHTML } from '../utils/EnhancedTableExamples';

export default {
  name: 'RichEditor',
  components: {
    Editor,
    Toolbar
  },
  props: {
    value: {
      type: String,
      default: ''
    }
  },
  data() {
    return {
      editor: null,
      editorHtml: '',
      mode: 'default', // 'default' 或 'simple'
      toolbarConfig: {
        toolbarKeys: [
          'bold', 'italic', 'underline', 'through', 'fontSize', 'fontFamily', 'color', 'bgColor',
          'insertLink', 'bulletedList', 'numberedList', 'todo', 'justifyLeft', 'justifyCenter', 'justifyRight',
          'insertTable', 'codeBlock', 'divider', 'undo', 'redo'
        ]
      },
      editorConfig: {
        placeholder: '请输入内容...',
        autoFocus: false,
        hoverbarKeys: {
          // 表格元素的悬浮菜单配置
          table: {
            menuKeys: [
              'tableHeader',
              'tableFullWidth',
              'insertTableRow',
              'deleteTableRow',
              'insertTableCol',
              'deleteTableCol',
              'deleteTable'
            ]
          }
        },
        MENU_CONF: {
          // 表格配置
          insertTable: {
            // 表格行列默认数量
            rowCount: 3,
            colCount: 3
          }
        }
      }
    }
  },
  watch: {
    value: {
      handler(newVal) {
        // 当外部传入的value改变时，更新编辑器内容
        if (this.editor && newVal !== this.editorHtml) {
          this.editorHtml = newVal;
        }
      },
      immediate: true
    }
  },
  mounted() {
    // 编辑器初始内容在handleCreated中设置
  },
  beforeDestroy() {
    // 销毁编辑器
    if (this.editor) {
      this.editor.destroy();
      this.editor = null;
    }
  },
  methods: {
    handleCreated(editor) {
      this.editor = editor;
      
      // 设置初始内容
      const initialContent = this.value || getAdvancedNestedTableHTML();
      this.editorHtml = initialContent;
      
      // 若初始时没有内容，editor.setHtml无效，需要一个setTimeout
      setTimeout(() => {
        if (!this.editor.isEmpty()) return;
        
        // 如果编辑器为空且有初始内容，设置内容
        if (initialContent) {
          this.editor.setHtml(initialContent);
        }
      }, 100);
    },
    
    handleChange(editor) {
      // 获取HTML并通过v-model同步
      const html = editor.getHtml();
      this.editorHtml = html;
      this.$emit('input', html);
    },
    
    getEditorContent() {
      return this.editor ? this.editor.getHtml() : '';
    },
    
    getEditorElement() {
      // 返回编辑器DOM元素，用于导出表格
      return this.editor ? this.editor.getEditableContainer() : null;
    },
    
    setContent(content) {
      console.log('开始设置内容:', content ? '有内容' : '无内容');
      
      // 设置编辑器内容
      if (!this.editor) {
        console.log('编辑器未初始化，延迟执行');
        // 如果编辑器未初始化，延迟执行
        setTimeout(() => {
          this.setContent(content);
        }, 200);
        return;
      }

      try {
        // 清空编辑器内容
        this.editor.clear();
        
        // 设置新内容
        if (content) {
          console.log('设置新内容');
          // 使用 DOM 操作的方式设置内容，避免解析错误
          const wrapDiv = document.createElement('div');
          wrapDiv.innerHTML = content;
          this.editor.setHtml(wrapDiv.innerHTML);
          this.editorHtml = wrapDiv.innerHTML;
        } else {
          // 如果内容为空，显示默认示例
          console.log('设置默认内容');
          const defaultContent = getAdvancedNestedTableHTML();
          const wrapDiv = document.createElement('div');
          wrapDiv.innerHTML = defaultContent;
          this.editor.setHtml(wrapDiv.innerHTML);
          this.editorHtml = wrapDiv.innerHTML;
        }
        
        // 更新状态
        this.$emit('input', this.editorHtml);
        console.log('内容设置完成', this.editor.getHtml().length);
        
      } catch (error) {
        console.error('设置编辑器内容失败:', error);
      }
    }
  }
}
</script>

<style>
.rich-editor {
  background-color: #fff;
}

.toolbar {
  border-bottom: 1px solid #f1f1f1;
}

.editor-content {
  height: 500px;
  overflow-y: auto;
}

.editor {
  height: 100%;
  overflow-y: hidden;
}

.w-e-text-container {
  height: 100% !important;
}
</style>