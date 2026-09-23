<script>
import { bitable, FieldType } from '@lark-base-open/js-sdk';
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
const EDITABLE_FIELD_TYPES = new Set([FieldType.Text]);

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

    const editing = ref(false);
    const editText = ref('');
    const saving = ref(false);
    const saveMsg = ref('');
    const fieldEditable = ref(false);

    let currentTableId = null;
    let currentFieldId = null;
    let currentRecordId = null;
    let cancelSelectionWatch = null;
    let fetchSeq = 0;
    let saveMsgTimer = null;

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

    function showSaveMsg(text, durationMs = 2000) {
      saveMsg.value = text;
      clearTimeout(saveMsgTimer);
      saveMsgTimer = setTimeout(() => { saveMsg.value = ''; }, durationMs);
    }

    function enterDemoMode() {
      demoMode.value = true;
      rawText.value = DEMO_MARKDOWN;
      tableName.value = 'Demo';
      fieldName.value = '预览';
      fieldEditable.value = true;
      loading.value = false;
    }

    const hasUnsavedChanges = computed(() => {
      return editing.value && editText.value !== rawText.value;
    });

    async function loadCell() {
      if (demoMode.value) return;

      const seq = ++fetchSeq;
      loading.value = true;
      errorMsg.value = '';
      noSelection.value = false;
      rawText.value = '';
      fieldName.value = '';
      tableName.value = '';
      editing.value = false;
      editText.value = '';
      fieldEditable.value = false;
      currentTableId = null;
      currentFieldId = null;
      currentRecordId = null;

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
        fieldEditable.value = EDITABLE_FIELD_TYPES.has(fieldMeta?.type);

        currentTableId = tableId;
        currentFieldId = fieldId;
        currentRecordId = recordId;

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

    function enterEdit() {
      editText.value = rawText.value;
      editing.value = true;
    }

    function exitEdit() {
      rawText.value = editText.value;
      editing.value = false;
    }

    async function saveCell() {
      if (demoMode.value) return;
      if (!currentTableId || !currentFieldId || !currentRecordId) return;

      saving.value = true;
      errorMsg.value = '';
      try {
        const table = await bitable.base.getTableById(currentTableId);
        await table.setCellValue(currentFieldId, currentRecordId, editText.value);
        rawText.value = editText.value;
        editing.value = false;
        showSaveMsg('已保存');
      } catch (err) {
        console.error('[MarkdownViewer] saveCell error:', err);
        errorMsg.value = err?.message || '保存失败';
      } finally {
        saving.value = false;
      }
    }

    function handleSelectionChange() {
      if (hasUnsavedChanges.value) {
        const discard = window.confirm('当前有未保存的修改，是否放弃？');
        if (!discard) return;
      }
      editing.value = false;
      loadCell();
    }

    const renderedHtml = computed(() => {
      if (!rawText.value) return '';
      return md.render(rawText.value);
    });

    const canEdit = computed(() => {
      return fieldEditable.value && !noSelection.value && !loading.value && !errorMsg.value;
    });

    const canSave = computed(() => {
      return !demoMode.value && !saving.value && editing.value;
    });

    onMounted(() => {
      loadCell();
      try {
        cancelSelectionWatch = bitable.base.onSelectionChange(() => {
          handleSelectionChange();
        });
      } catch (e) {
        console.warn('[MarkdownViewer] onSelectionChange 注册失败:', e);
      }
    });

    onUnmounted(() => {
      fetchSeq++;
      clearTimeout(saveMsgTimer);
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
      editing,
      editText,
      saving,
      saveMsg,
      fieldEditable,
      renderedHtml,
      canEdit,
      canSave,
      enterEdit,
      exitEdit,
      saveCell,
    };
  },
};
</script>

<template>
  <div class="md-viewer">
    <!-- header -->
    <div class="md-viewer__header">
      <div class="md-viewer__header-left">
        <template v-if="tableName || fieldName">
          <span class="md-viewer__badge" v-if="tableName">{{ tableName }}</span>
          <span class="md-viewer__sep" v-if="tableName && fieldName">/</span>
          <span class="md-viewer__badge md-viewer__badge--field" v-if="fieldName">{{ fieldName }}</span>
        </template>
      </div>
      <div class="md-viewer__header-right">
        <span v-if="saveMsg" class="md-viewer__toast">{{ saveMsg }}</span>
        <template v-if="editing">
          <button
            class="md-viewer__btn md-viewer__btn--save"
            :disabled="!canSave"
            @click="saveCell"
          >{{ saving ? '保存中…' : '保存' }}</button>
          <button class="md-viewer__btn" @click="exitEdit">预览</button>
        </template>
        <template v-else>
          <button
            class="md-viewer__btn"
            :disabled="!canEdit"
            :title="!fieldEditable && !demoMode ? '仅文本字段支持编辑' : (demoMode ? '编辑（Demo 模式不可保存）' : '编辑')"
            @click="enterEdit"
          >编辑</button>
        </template>
      </div>
    </div>

    <!-- states -->
    <div v-if="loading" class="md-viewer__status">加载中…</div>
    <div v-else-if="errorMsg && !editing" class="md-viewer__status md-viewer__status--error">{{ errorMsg }}</div>
    <div v-else-if="noSelection" class="md-viewer__status">请在多维表格中选中一个单元格</div>

    <!-- editing -->
    <template v-else-if="editing">
      <div v-if="errorMsg" class="md-viewer__inline-error">{{ errorMsg }}</div>
      <textarea
        class="md-viewer__editor"
        v-model="editText"
        placeholder="输入 Markdown 内容…"
        spellcheck="false"
      ></textarea>
    </template>

    <!-- empty -->
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
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 6px 12px;
  font-size: 12px;
  color: #646a73;
  border-bottom: 1px solid #e5e6eb;
  background: #fafafa;
  gap: 8px;
}
.md-viewer__header-left {
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  min-width: 0;
}
.md-viewer__header-right {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  gap: 6px;
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
.md-viewer__btn {
  padding: 3px 10px;
  font-size: 12px;
  border: 1px solid #c9cdd4;
  border-radius: 4px;
  background: #fff;
  color: #1f2329;
  cursor: pointer;
  white-space: nowrap;
  transition: background 0.15s, border-color 0.15s;
}
.md-viewer__btn:hover:not(:disabled) {
  background: #f2f3f5;
  border-color: #8f959e;
}
.md-viewer__btn:disabled {
  opacity: 0.45;
  cursor: not-allowed;
}
.md-viewer__btn--save {
  background: #3370ff;
  color: #fff;
  border-color: #3370ff;
}
.md-viewer__btn--save:hover:not(:disabled) {
  background: #245bdb;
  border-color: #245bdb;
}
.md-viewer__toast {
  color: #00b42a;
  font-size: 12px;
  animation: fadeIn 0.2s;
}
@keyframes fadeIn {
  from { opacity: 0; }
  to   { opacity: 1; }
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
.md-viewer__inline-error {
  padding: 6px 16px;
  font-size: 12px;
  color: #f53f3f;
  background: #fff2f0;
  border-bottom: 1px solid #fde2e2;
}
.md-viewer__editor {
  flex: 1;
  width: 100%;
  padding: 12px 16px;
  font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
  font-size: 13px;
  line-height: 1.6;
  border: none;
  outline: none;
  resize: none;
  background: #fff;
  color: #1f2329;
  tab-size: 2;
}
.md-viewer__editor:focus {
  background: #fafbfc;
}
.md-viewer__body {
  flex: 1;
  overflow: auto;
  padding: 12px 16px;
}
</style>
