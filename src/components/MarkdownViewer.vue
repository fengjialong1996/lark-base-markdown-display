<script>
import { bitable } from '@lark-base-open/js-sdk';
import { ref, computed, onMounted, onUnmounted } from 'vue';
import md from '../lib/markdown.js';

export default {
  name: 'MarkdownViewer',
  setup() {
    const rawText = ref('');
    const fieldName = ref('');
    const tableName = ref('');
    const loading = ref(true);
    const errorMsg = ref('');
    const noSelection = ref(false);

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

    async function loadCell() {
      const seq = ++fetchSeq;
      loading.value = true;
      errorMsg.value = '';
      noSelection.value = false;
      rawText.value = '';
      fieldName.value = '';
      tableName.value = '';

      try {
        const selection = await bitable.base.getSelection();

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
      cancelSelectionWatch = bitable.base.onSelectionChange(() => {
        loadCell();
      });
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
