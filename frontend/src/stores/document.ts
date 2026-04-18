import { defineStore } from "pinia";
import { ref, watch, type Ref } from "vue";

import type { ProjectAssetItem } from "@/api/documents";

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
        assets: Array.isArray(parsed.assets) ? parsed.assets : [],
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
  assets: ProjectAssetItem[];
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
  const assets: Ref<ProjectAssetItem[]> = ref(saved?.assets ?? []);

  watch(
    [projectName, markdown, assets],
    () => {
      saveState({
        projectName: projectName.value,
        markdown: markdown.value,
        assets: assets.value,
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

  return {
    projectName,
    markdown,
    assets,
    setProjectName,
    setMarkdown,
    setAssets,
    upsertAsset,
    removeAsset,
  };
});
