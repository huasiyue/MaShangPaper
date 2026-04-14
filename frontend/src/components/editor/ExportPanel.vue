<script setup lang="ts">
const props = defineProps<{
  convertLoading: boolean;
  reviewLoading: boolean;
  formatLoading: boolean;
  exportProjectLoading: boolean;
  markdownDisabled: boolean;
  wordDisabled: boolean;
}>();

const emit = defineEmits<{
  (e: "export-project"): void;
  (e: "convert"): void;
  (e: "review"): void;
  (e: "format"): void;
}>();
</script>

<template>
  <a-space wrap size="small">
    <a-button size="small" :loading="props.exportProjectLoading" :disabled="props.markdownDisabled" @click="emit('export-project')">
      项目包
    </a-button>
    <a-button type="primary" size="small" :loading="props.convertLoading" :disabled="props.markdownDisabled" @click="emit('convert')">
      导出
    </a-button>
    <a-button size="small" :loading="props.reviewLoading" :disabled="props.wordDisabled" @click="emit('review')">
      审查
    </a-button>
    <a-button type="outline" size="small" :loading="props.formatLoading" :disabled="props.wordDisabled" @click="emit('format')">
      格式化
    </a-button>
  </a-space>
</template>
