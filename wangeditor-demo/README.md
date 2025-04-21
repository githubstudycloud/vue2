# WangEditor 富文本编辑器与 Excel 导出功能

这是一个基于 Vue 2.7 和 WangEditor 的富文本编辑器项目，实现了复杂表格编辑和导出为Excel的功能。

## 项目结构

```
wangeditor-demo/
├── public/             
│   └── index.html      # HTML入口文件
├── src/                
│   ├── components/     
│   │   └── RichEditor.vue  # 富文本编辑器组件
│   ├── utils/         
│   │   ├── ExcelExporter.js           # 基础Excel导出工具
│   │   ├── ExcelExporterPro.js        # 增强版Excel导出工具
│   │   ├── ExcelJSExporter.js         # 使用ExcelJS的导出工具
│   │   ├── ExcelJSExporterComplete.js # 完整样式版导出工具
│   │   ├── ExcelJSExporterFixed.js    # 修复数据丢失版导出工具 (主文件)
│   │   ├── excel/                     # Excel工具辅助模块文件夹
│   │   │   ├── ExcelStyleHelper.js    # Excel样式处理辅助工具
│   │   │   └── ExcelTableHelper.js    # Excel表格内容处理辅助工具
│   │   ├── TableExamples.js           # 表格示例数据
│   │   └── EnhancedTableExamples.js   # 增强版表格示例数据
│   ├── App.vue         # 主应用组件
│   └── main.js         # 应用入口文件
├── package.json        # 项目配置文件
├── vue.config.js       # Vue配置文件
└── README.md           # 项目说明文档
```

## 技术栈

- Vue 2.7
- @wangeditor/editor: 富文本编辑器核心
- @wangeditor/editor-for-vue: Vue集成包
- xlsx: Excel处理库
- exceljs: 增强的Excel处理库（支持更多样式）
- file-saver: 文件保存工具

## 功能特性

1. **富文本编辑**：
   - 支持多种文本格式（粗体、斜体、下划线等）
   - 支持文字颜色、背景色、字体大小等设置
   - 支持文本对齐（左对齐、居中、右对齐）

2. **表格功能**：
   - 支持创建复杂表格
   - 支持合并单元格（rowspan/colspan）
   - 支持多层嵌套表格
   - 支持各种表格样式（边框、背景色等）

3. **Excel导出**：
   - 基础导出功能（ExcelExporter.js）
   - 带样式导出（ExcelJSExporter.js）
   - 完整样式支持（ExcelJSExporterComplete.js）
   - 数据完整性修复版（ExcelJSExporterFixed.js）
   - 重构兼容版（ExcelJSExporterFixed.js + excel/目录下的辅助工具）

## 最新改进说明

最新版本导出工具进行了重大改进，主要包括：

1. **代码结构优化**：
   - 将原 `ExcelJSExporterFixed.js` 拆分为3个独立模块：
     - `ExcelJSExporterFixed.js` - 主文件，负责导出流程控制
     - `excel/ExcelStyleHelper.js` - 负责样式处理
     - `excel/ExcelTableHelper.js` - 负责表格内容和结构处理
   - 函数拆分和重组，提高代码可读性和可维护性

2. **兼容性增强**：
   - 移除了可选链操作符（`?.`）等新语法，兼容Babel 5.x等旧版本编译环境
   - 使用传统函数表达式代替箭头函数
   - 增加各种边界情况检查，提高代码健壮性

3. **圈复杂度优化**：
   - 每个函数的圈复杂度控制在15以下
   - 抽取公共逻辑到辅助函数
   - 单一职责原则，每个函数只做一件事

4. **表格渲染增强**：
   - 同时支持v-if和v-show隐藏的表格导出
   - 通过临时显示隐藏表格的方式处理v-show表格
   - 增加表格预处理函数，修复不完整表格结构

5. **样式呈现优化**：
   - 改进列宽自适应算法，根据内容类型智能计算宽度
   - 更好地支持中文和混合文本内容的宽度计算
   - 限制过长列的宽度，防止表格过宽

## 使用方法

使用最新的Excel导出功能非常简单：

```javascript
// 在Vue组件中引入
import ExcelJSExporterFixed from './utils/ExcelJSExporterFixed';

// 调用导出方法
const editorElement = this.$refs.editor.getEditorElement();
const success = await ExcelJSExporterFixed.exportTableToExcel(editorElement);
```

## 关键实现细节

### 1. 表格样式处理

- 通过检测内联样式、计算样式和元素标签来识别文本格式
- 使用递归方式处理嵌套表格
- 支持将RGB颜色转换为Excel的ARGB格式
- 支持字体、对齐、背景色等样式完整保留

### 2. 合并单元格处理

- 使用skipCells机制跟踪被合并的单元格
- 正确处理rowspan和colspan属性
- 确保导出后的Excel保持相同的单元格结构

### 3. 嵌套表格处理

- 在导出时将嵌套表格转换为文本格式
- 使用分隔符保留嵌套表格结构信息
- 避免数据丢失问题

### 4. V-if和V-show的表格支持

- 处理v-show的表格：临时设置display为table，导出后恢复
- 处理v-if的表格：从整个文档范围查找表格元素
- 使用特殊标记防止重复处理同一表格

## 开发注意事项

1. **样式检测优先级**：
   - 优先使用内联样式（style属性）
   - 其次检查特定HTML元素（如`<strong>`、`<em>`等）
   - 最后使用计算样式

2. **性能考虑**：
   - 对于大型表格，需要注意性能优化
   - 避免过多的DOM操作和样式计算
   - 使用适当的缓存机制

3. **兼容性**：
   - 确保在不同浏览器中的表现一致
   - 处理某些Excel不支持的样式格式
   - 考虑不同版本Excel的兼容性
   - 代码语法兼容旧版本Babel

## 已知问题

1. 某些特殊字符可能无法正确导出
2. 极复杂的嵌套表格可能导致性能问题
3. 部分Web样式无法完美映射到Excel

## 未来改进方向

1. 添加更多表格模板
2. 支持更多Excel样式（如条件格式）
3. 优化大数据量的导出性能
4. 支持导入Excel文件到编辑器
5. 添加导出进度指示
6. 提供更多导出自定义选项
