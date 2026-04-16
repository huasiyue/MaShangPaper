import { defineStore } from "pinia";
import { ref, watch, type Ref } from "vue";

import type { ProjectAssetItem, ReviewResponse, SchoolId, ThesisType } from "@/api/documents";

const STORAGE_KEY = "mashangpaper-document-store";

const DEFAULT_MARKDOWN = `# 题目

## 摘要

请在这里输入论文摘要。

## 第一章 绪论

在这里开始撰写正文。`;

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      return {
        projectName: parsed.projectName ?? "paper-project",
        markdown: parsed.markdown ?? DEFAULT_MARKDOWN,
        schoolId: parsed.schoolId ?? "sdfmu",
        thesisType: parsed.thesisType ?? "thesis",
        assets: Array.isArray(parsed.assets) ? parsed.assets : [],
        reviewResult: parsed.reviewResult ?? null,
      };
    }
  } catch {
    // ignore
  }
  return null;
}

function saveState(state: {
  projectName: string;
  markdown: string;
  schoolId: SchoolId;
  thesisType: ThesisType;
  assets: ProjectAssetItem[];
  reviewResult: ReviewResponse | null;
}) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // ignore
  }
}

export const useDocumentStore = defineStore("document", () => {
  const saved = loadState();

  const projectName: Ref<string> = ref(saved?.projectName ?? "paper-project");
  const markdown: Ref<string> = ref(saved?.markdown ?? DEFAULT_MARKDOWN);
  const schoolId: Ref<SchoolId> = ref(saved?.schoolId ?? "sdfmu");
  const thesisType: Ref<ThesisType> = ref(saved?.thesisType ?? "thesis");
  const assets: Ref<ProjectAssetItem[]> = ref(saved?.assets ?? []);
  const reviewResult: Ref<ReviewResponse | null> = ref(saved?.reviewResult ?? null);

  watch(
    [projectName, markdown, schoolId, thesisType, assets, reviewResult],
    () => {
      saveState({
        projectName: projectName.value,
        markdown: markdown.value,
        schoolId: schoolId.value,
        thesisType: thesisType.value,
        assets: assets.value,
        reviewResult: reviewResult.value,
      });
    },
    { deep: true }
  );

  function setProjectName(value: string) {
    projectName.value = value;
  }
  function setMarkdown(value: string) {
    markdown.value = value;
  }
  function setSchoolId(value: SchoolId) {
    schoolId.value = value;
  }
  function setThesisType(value: ThesisType) {
    thesisType.value = value;
  }
  function setAssets(value: ProjectAssetItem[]) {
    assets.value = value;
  }
  function upsertAsset(value: ProjectAssetItem) {
    const index = assets.value.findIndex((item) => item.asset_id === value.asset_id);
    if (index === -1) {
      assets.value = [...assets.value, value];
      return;
    }
    const next = [...assets.value];
    next[index] = value;
    assets.value = next;
  }
  function removeAsset(assetId: string) {
    assets.value = assets.value.filter((item) => item.asset_id !== assetId);
  }
  function setReviewResult(value: ReviewResponse | null) {
    reviewResult.value = value;
  }

  return {
    projectName,
    markdown,
    schoolId,
    thesisType,
    assets,
    reviewResult,
    setProjectName,
    setMarkdown,
    setSchoolId,
    setThesisType,
    setAssets,
    upsertAsset,
    removeAsset,
    setReviewResult,
  };
});

