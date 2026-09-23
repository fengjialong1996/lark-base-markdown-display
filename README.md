# Lark Base Markdown Display

飞书多维表格插件 —— 将选中单元格的文本内容以 Markdown 格式渲染预览。

## 功能

- **跟随选区**：自动监听当前选中的单元格，切换单元格时实时刷新渲染。
- **Markdown 全特性**：标题、列表、表格、引用、链接、图片、水平线。
- **围栏代码高亮**：支持 SQL、YAML、JSON、Go、Python、JavaScript 等语言的语法着色。
- **原生 HTML**：允许单元格中嵌入 HTML 标签直接渲染（仅适用于可信内容）。

## 开发

```bash
npm install
npm run dev
```

## 构建与发布

```bash
npm run build
```

构建产物在 `dist/` 目录，连同该目录提交后填写发布表单：
- [共享表单（中文）](https://feishu.feishu.cn/share/base/form/shrcnGFgOOsFGew3SDZHPhzkM0e)

## 安全提示

本插件启用了 Markdown 原生 HTML 渲染（`html: true`），单元格中的 `<script>`、`<iframe>` 等标签会被浏览器执行。请仅在可信数据表中使用。如需禁用，修改 `src/lib/markdown.js` 中 `html: false`。

## 参考

- [多维表格插件开发指南](https://bytedance.feishu.cn/docx/HazFdSHH9ofRGKx8424cwzLlnZc)
- [Base JS SDK API](https://bytedance.feishu.cn/docx/HjCEd1sPzoVnxIxF3LrcKnepnUf)
