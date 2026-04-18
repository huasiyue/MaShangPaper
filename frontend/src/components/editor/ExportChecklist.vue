<script setup lang="ts">
import { computed } from "vue";


const props = defineProps<{
  markdown: string;
  assetCount: number;
}>();

const emit = defineEmits<{
  (e: "insert", snippet: string): void;
}>();

function hasHeading(text: string, ...candidates: string[]) {
  const normalized = text.toLowerCase().replace(/\s+/g, "");
  return candidates.some((candidate) => normalized.includes(`#${candidate.toLowerCase().replace(/\s+/g, "")}`));
}

const checks = computed(() => {
  const markdown = props.markdown;
  const rows = [
    { label: "封面", ok: hasHeading(markdown, "封面") },
    { label: "中文摘要", ok: hasHeading(markdown, "摘要") },
    { label: "英文摘要", ok: hasHeading(markdown, "abstract") },
    { label: "目录", ok: hasHeading(markdown, "目录") },
    { label: "参考文献", ok: hasHeading(markdown, "参考文献") },
    { label: "致谢", ok: hasHeading(markdown, "致谢") },
  ];
  return rows;
});

const missingLabels = computed(() =>
  checks.value.filter((item) => !item.ok).map((item) => item.label),
);

const imageCount = computed(() => (props.markdown.match(/!\[(.*?)\]\((.*?)\)/g) ?? []).length);
const tableCount = computed(() => (props.markdown.match(/^\|.+\|\s*$/gm) ?? []).length / 2);
const referenceCount = computed(() => (props.markdown.match(/^\[\d+\]/gm) ?? []).length);
const charCount = computed(() => props.markdown.replace(/\s+/g, "").length);

const readiness = computed(() => {
  if (missingLabels.value.length === 0 && referenceCount.value >= 15) {
    return {
      label: "可导出",
      color: "green",
      description: "主要结构已具备，适合直接生成 Word。",
    };
  }
  return {
    label: "建议补充",
    color: "orange",
    description:
      missingLabels.value.length > 0
        ? `建议先补齐：${missingLabels.value.join("、")}`
        : "建议先补充参考文献到 15 条以上。",
  };
});

const quickInsertItems = [
  { label: "封面", snippet: "# 封面\n题目：\n教学机构：医学信息与人工智能学院\n专业：\n年级、班级：\n学号：\n学生姓名：\n指导教师：\n企业导师：普通本科删除该行\n完成日期：\n" },
  { label: "摘要", snippet: "# 摘要\n这里写中文摘要。\n\n关键词：关键词一；关键词二；关键词三\n" },
  { label: "Abstract", snippet: "# Abstract\nWrite abstract here.\n\nKey words: keyword1; keyword2; keyword3\n" },
  { label: "参考文献", snippet: "# 参考文献\n[1] 作者. 题名[J]. 期刊名, 2024, 1(1): 1-10.\n" },
  { label: "致谢", snippet: "# 致谢\n感谢指导教师、同学和家人的支持。\n" },
];
</script>

<template>
  <a-card class="checklist-card panel-shell" :bordered="false">
    <template #title>导出检查</template>
    <template #extra>
      <a-tag :color="readiness.color" bordered>{{ readiness.label }}</a-tag>
    </template>

    <a-space direction="vertical" fill size="large">
      <a-alert :type="readiness.color === 'green' ? 'success' : 'warning'" :show-icon="true">
        {{ readiness.description }}
      </a-alert>

      <div class="stats-grid">
        <div class="stats-tile">
          <div class="stats-tile__label">字数</div>
          <div class="stats-tile__value">{{ charCount }}</div>
        </div>
        <div class="stats-tile">
          <div class="stats-tile__label">图片</div>
          <div class="stats-tile__value">{{ Math.max(imageCount, props.assetCount) }}</div>
        </div>
        <div class="stats-tile">
          <div class="stats-tile__label">表格</div>
          <div class="stats-tile__value">{{ Math.floor(tableCount) }}</div>
        </div>
        <div class="stats-tile">
          <div class="stats-tile__label">参考文献</div>
          <div class="stats-tile__value">{{ referenceCount }}</div>
        </div>
      </div>

      <div class="checklist-items">
        <div v-for="item in checks" :key="item.label" class="checklist-item">
          <span>{{ item.label }}</span>
          <a-tag :color="item.ok ? 'green' : 'orange'" bordered>
            {{ item.ok ? "已包含" : "缺失" }}
          </a-tag>
        </div>
      </div>

      <div class="quick-actions">
        <div class="quick-actions__title">快捷插入</div>
        <a-space wrap size="small">
          <a-button
            v-for="item in quickInsertItems"
            :key="item.label"
            size="mini"
            @click="emit('insert', item.snippet)"
          >
            {{ item.label }}
          </a-button>
        </a-space>
      </div>
    </a-space>
  </a-card>
</template>

<style scoped>
.checklist-card {
  min-height: 360px;
  border-radius: 28px;
}

.stats-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 12px;
}

.stats-tile {
  padding: 14px 16px;
  border-radius: 18px;
  background: rgba(255, 255, 255, 0.82);
  border: 1px solid rgba(22, 93, 255, 0.08);
}

.stats-tile__label {
  font-size: 12px;
  color: var(--msp-text-muted);
}

.stats-tile__value {
  margin-top: 6px;
  font-size: 24px;
  font-weight: 700;
  color: var(--msp-text);
}

.checklist-items {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.checklist-item {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding-bottom: 8px;
  border-bottom: 1px dashed rgba(15, 23, 42, 0.08);
}

.quick-actions {
  padding-top: 4px;
}

.quick-actions__title {
  margin-bottom: 8px;
  font-size: 12px;
  color: var(--msp-text-muted);
}
</style>
