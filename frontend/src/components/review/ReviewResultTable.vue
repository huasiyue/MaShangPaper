<script setup lang="ts">
import { computed } from "vue";

import type { ReviewResponse } from "@/api/documents";


const props = defineProps<{
  review: ReviewResponse | null;
}>();

const rows = computed(() =>
  (props.review?.issues ?? []).map((issue, index) => ({
    ...issue,
    key: `${issue.level}-${issue.location}-${index}`,
  })),
);

function levelColor(level: string): string {
  if (level.includes("错误")) {
    return "#f53f3f";
  }
  if (level.includes("警告")) {
    return "#ff7d00";
  }
  return "#165dff";
}
</script>

<template>
  <a-card class="review-card panel-shell" :bordered="false">
    <template #title>结果</template>

    <template v-if="props.review">
      <a-space direction="vertical" fill size="large">
        <a-space wrap>
          <a-tag color="red" bordered>错误 {{ props.review.error_count }}</a-tag>
          <a-tag color="orange" bordered>警告 {{ props.review.warning_count }}</a-tag>
          <a-tag color="arcoblue" bordered>提示 {{ props.review.info_count }}</a-tag>
          <a-tag bordered>总计 {{ props.review.total_issues }}</a-tag>
        </a-space>

        <a-table :data="rows" :pagination="false" row-key="key">
          <a-table-column title="级别" data-index="level">
            <template #cell="{ record }">
              <a-tag :color="levelColor(record.level)">{{ record.level }}</a-tag>
            </template>
          </a-table-column>
          <a-table-column title="位置" data-index="location" />
          <a-table-column title="说明" data-index="description" />
          <a-table-column title="建议" data-index="suggestion" />
        </a-table>

        <div class="review-report">
          <pre>{{ props.review.report_text }}</pre>
        </div>
      </a-space>
    </template>

    <template v-else>
      <div class="review-empty">
        <a-empty description="暂无结果" />
      </div>
    </template>
  </a-card>
</template>

<style scoped>
.review-card {
  min-height: 320px;
  border-radius: 28px;
}

.review-empty {
  min-height: 220px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.review-report {
  background: rgba(255, 255, 255, 0.76);
  border: 1px solid rgba(22, 93, 255, 0.08);
  border-radius: 18px;
  padding: 16px;
}

.review-report pre {
  margin: 0;
  white-space: pre-wrap;
  word-break: break-word;
  font-family: "SFMono-Regular", "JetBrains Mono", monospace;
  font-size: 12px;
  line-height: 1.6;
}
</style>
