<script>
import { bitable } from '@lark-base-open/js-sdk';
import { ref, computed, onMounted, onUnmounted } from 'vue';
import md from '../lib/markdown.js';

const DEMO_MARKDOWN = `# Markdown 渲染预览（Demo 模式）

> 当前在本地浏览器中运行，未连接飞书多维表格 SDK。
> 请在飞书多维表格中通过"添加插件"加载本插件。

---

## 支持的 Markdown 特性

### 文本格式
**粗体** / *斜体* / ~~删除线~~ / \`行内代码\`

### 列表
- 无序列表项 A
  - 嵌套项 A-1
  - 嵌套项 A-2
- 无序列表项 B

1. 有序列表第一项
2. 有序列表第二项

### 链接与图片
[飞书多维表格开发文档](https://bytedance.feishu.cn/docx/HazFdSHH9ofRGKx8424cwzLlnZc)

### 表格
| 层级 | 名称 | 说明 |
|------|------|------|
| L0 | Entity | 发行人 / 主体 |
| L1 | Instrument | 金融工具 |
| L2 | Tradable | 可交易对象 |

### 代码高亮

\`\`\`sql
SELECT name, last_source_change_id
FROM entity_attribute
WHERE entity_id = 42;
\`\`\`

\`\`\`yaml
assert:
  row_count: 1
  fields:
    name:
      equals: "ACME Corp"
\`\`\`

\`\`\`javascript
const selection = await bitable.base.getSelection();
console.log(selection.tableId, selection.fieldId);
\`\`\`

### 引用块
> 这是一段引用文本，可以嵌套 **粗体** 和 \`代码\`。
`;

const SDK_TIMEOUT_MS = 3000;

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error('__SDK_TIMEOUT__')), ms)
    ),
  ]);
}

export default {
  name: 'MarkdownViewer',
  setup() {
    const rawText = ref('');
    const fieldName = ref('');
    const tableName = ref('');
    const loading = ref(true);
    const errorMsg = ref('');
    const noSelection = ref(false);
    const demoMode = ref(false);

    let cancelSelectionWatch = null;
    let fetchSeq = 0;

    function extractCellText(cellValue) {
      if (cellValue == null) return '';
      if (typeof cellValue === 'string') return cellValue;
      if (typeof cellValue === 'number' || typeof cellValue === 'boolean') return String(cellValue);
      if (Array.isArray(cellValue)) {
        return cellValue
          .map((seg) => {
            if (typeof seg === 'string') return seg;
            if (seg && typeof seg.text === 'string') return seg.text;
            if (seg && typeof seg.link === 'string') return seg.link;
            if (seg && typeof seg.name === 'string') return seg.name;
            return '';
          })
          .join('');
      }
      if (typeof cellValue === 'object' && cellValue.text) return cellValue.text;
      try { return JSON.stringify(cellValue, null, 2); } catch { return String(cellValue); }
    }

    function enterDemoMode() {
      demoMode.value = true;
      rawText.value = DEMO_MARKDOWN;
      tableName.value = 'Demo';
      fieldName.value = '预览';
      loading.value = false;
    }

    async function loadCell() {
      if (demoMode.value) return;

      const seq = ++fetchSeq;
      loading.value = true;
      errorMsg.value = '';
      noSelection.value = false;
      rawText.value = '';
      fieldName.value = '';
      tableName.value = '';

      try {
        const selection = await withTimeout(
          bitable.base.getSelection(),
          SDK_TIMEOUT_MS
        );

        if (seq !== fetchSeq) return;

        const { tableId, fieldId, recordId } = selection || {};
        if (!tableId || !fieldId || !recordId) {
          noSelection.value = true;
          loading.value = false;
          return;
        }

        const table = await bitable.base.getTableById(tableId);
        if (seq !== fetchSeq) return;

        const [fieldMeta, tableMeta] = await Promise.all([
          table.getFieldMetaById(fieldId),
          bitable.base.getTableMetaList().then((list) =>
            list.find((t) => t.id === tableId)
          ),
        ]);
        if (seq !== fetchSeq) return;

        fieldName.value = fieldMeta?.name || fieldId;
        tableName.value = tableMeta?.name || tableId;

        const cellValue = await table.getCellValue(fieldId, recordId);
        if (seq !== fetchSeq) return;

        rawText.value = extractCellText(cellValue);
      } catch (err) {
        if (seq !== fetchSeq) return;
        if (err?.message === '__SDK_TIMEOUT__') {
          console.warn('[MarkdownViewer] SDK 超时，进入 Demo 模式');
          enterDemoMode();
          return;
        }
        console.error('[MarkdownViewer] loadCell error:', err);
        errorMsg.value = err?.message || '读取单元格失败';
      } finally {
        if (seq === fetchSeq) loading.value = false;
      }
    }

    const renderedHtml = computed(() => {
      if (!rawText.value) return '';
      return md.render(rawText.value);
    });

    onMounted(() => {
      loadCell();
      try {
        cancelSelectionWatch = bitable.base.onSelectionChange(() => {
          loadCell();
        });
      } catch (e) {
        console.warn('[MarkdownViewer] onSelectionChange 注册失败:', e);
      }
    });

    onUnmounted(() => {
      fetchSeq++;
      if (typeof cancelSelectionWatch === 'function') {
        cancelSelectionWatch();
        cancelSelectionWatch = null;
      }
    });

    return {
      rawText,
      fieldName,
      tableName,
      loading,
      errorMsg,
      noSelection,
      demoMode,
      renderedHtml,
    };
  },
};
</script>

<template>
  <div class="md-viewer">
    <!-- header -->
    <div v-if="tableName || fieldName" class="md-viewer__header">
      <span class="md-viewer__badge" v-if="tableName">{{ tableName }}</span>
      <span class="md-viewer__sep" v-if="tableName && fieldName">/</span>
      <span class="md-viewer__badge md-viewer__badge--field" v-if="fieldName">{{ fieldName }}</span>
    </div>

    <!-- states -->
    <div v-if="loading" class="md-viewer__status">加载中…</div>
    <div v-else-if="errorMsg" class="md-viewer__status md-viewer__status--error">{{ errorMsg }}</div>
    <div v-else-if="noSelection" class="md-viewer__status">请在多维表格中选中一个单元格</div>
    <div v-else-if="!rawText" class="md-viewer__status">单元格为空</div>

    <!-- rendered markdown -->
    <div v-else class="md-viewer__body markdown-body" v-html="renderedHtml"></div>
  </div>
</template>

<style scoped>
.md-viewer {
  display: flex;
  flex-direction: column;
  height: 100%;
  overflow: hidden;
}
.md-viewer__header {
  flex-shrink: 0;
  padding: 8px 12px;
  font-size: 12px;
  color: #646a73;
  border-bottom: 1px solid #e5e6eb;
  background: #fafafa;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.md-viewer__badge {
  display: inline-block;
  padding: 1px 6px;
  border-radius: 4px;
  background: #f0f1f5;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}
.md-viewer__badge--field {
  background: #e8f3ff;
  color: #3370ff;
}
.md-viewer__sep {
  margin: 0 4px;
  color: #c9cdd4;
}
.md-viewer__status {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #8f959e;
  font-size: 14px;
  padding: 24px;
  text-align: center;
}
.md-viewer__status--error {
  color: #f53f3f;
}
.md-viewer__body {
  flex: 1;
  overflow: auto;
  padding: 12px 16px;
}
</style>
