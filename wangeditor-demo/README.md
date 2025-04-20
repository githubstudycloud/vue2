# WangEditor 富文本编辑器示例

这是一个基于Vue 2.7和@wangeditor/editor的富文本编辑器示例项目，展示了多层嵌套表格的使用以及表格导出为Excel的功能。

## 功能特点

1. 使用Vue 2.7框架
2. 集成@wangeditor/editor富文本编辑器
3. 支持多层嵌套表格
4. 支持多种文本格式（不同字体、大小、颜色、对齐方式等）
5. 支持将编辑器中的表格导出为Excel格式

## 安装与运行

1. 确保已安装Node.js环境
2. 双击`install.bat`脚本自动安装依赖并启动项目
3. 或者手动执行以下命令：
   ```bash
   npm install
   npm run serve
   ```
4. 浏览器访问 http://localhost:8080 查看效果

## 项目结构

```
wangeditor-demo/
├── public/             # 静态资源文件夹
│   └── index.html      # HTML入口文件
├── src/                # 源代码文件夹
│   ├── components/     # 组件文件夹
│   │   └── RichEditor.vue  # 富文本编辑器组件
│   ├── App.vue         # 主应用组件
│   └── main.js         # 应用入口文件
├── package.json        # 项目配置文件
├── vue.config.js       # Vue配置文件
└── README.md           # 项目说明文档
```

## 主要依赖

- Vue 2.7
- @wangeditor/editor: 富文本编辑器核心
- @wangeditor/editor-for-vue: Vue集成包
- xlsx: Excel处理库
- file-saver: 文件保存工具

## 使用说明

1. 页面加载时会自动创建包含示例嵌套表格的编辑器
2. 可以使用编辑器工具栏修改文本格式和表格内容
3. 点击"导出为Excel"按钮可将编辑器中的表格导出为Excel文件