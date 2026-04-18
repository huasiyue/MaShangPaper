<script setup lang="ts">
const props = defineProps<{
  convertLoading: boolean;
  exportProjectLoading: boolean;
  markdownDisabled: boolean;
}>();

const emit = defineEmits<{
  (e: "export-project"): void;
  (e: "convert"): void;
}>();
</script>

<template>
  <a-space wrap size="small">
    <a-popconfirm
      content="确认导出当前 Markdown 与素材为项目包吗？"
      ok-text="确认导出"
      cancel-text="取消"
      @ok="emit('export-project')"
    >
      <a-button size="small" :loading="props.exportProjectLoading" :disabled="props.markdownDisabled">
        导出项目包
      </a-button>
    </a-popconfirm>
    <a-popconfirm
      content="确认导出当前 Markdown 为 Word 吗？"
      ok-text="确认导出"
      cancel-text="取消"
      @ok="emit('convert')"
    >
      <a-button type="primary" size="small" :loading="props.convertLoading" :disabled="props.markdownDisabled">
        导出 Word
      </a-button>
    </a-popconfirm>
  </a-space>
</template>
