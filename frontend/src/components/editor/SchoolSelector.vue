<script setup lang="ts">
import { computed } from "vue";

import type { SchoolId, ThesisType } from "@/api/documents";


const props = defineProps<{
  schoolId: SchoolId;
  thesisType: ThesisType;
}>();

const emit = defineEmits<{
  (e: "update:schoolId", value: SchoolId): void;
  (e: "update:thesisType", value: ThesisType): void;
}>();

const schoolProxy = computed({
  get: () => props.schoolId,
  set: (value: SchoolId) => emit("update:schoolId", value),
});

const thesisProxy = computed({
  get: () => props.thesisType,
  set: (value: ThesisType) => emit("update:thesisType", value),
});
</script>

<template>
  <div class="school-selector">
    <div class="selector-block">
      <span class="selector-label">学校</span>
      <a-select v-model="schoolProxy" size="small" style="width: 260px">
        <a-option value="sdfmu">山东第一医科大学（山东省医学科学院）</a-option>
        <a-option value="yzu">扬州大学</a-option>
      </a-select>
    </div>

    <div class="selector-block">
      <span class="selector-label">类型</span>
      <a-radio-group v-model="thesisProxy" type="button" size="small">
        <a-radio value="thesis">论文</a-radio>
        <a-radio value="design_report">设计</a-radio>
      </a-radio-group>
    </div>
  </div>
</template>

<style scoped>
.school-selector {
  display: flex;
  align-items: center;
  gap: 14px;
  flex-wrap: wrap;
}

.selector-block {
  display: flex;
  align-items: center;
  gap: 8px;
}

.selector-label {
  font-size: 12px;
  color: var(--msp-text-muted);
}
</style>
