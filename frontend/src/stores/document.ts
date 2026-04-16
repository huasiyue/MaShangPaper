import { defineStore } from "pinia";

import type { ProjectAssetItem, ReviewResponse, SchoolId, ThesisType } from "@/api/documents";


interface DocumentState {
  projectName: string;
  markdown: string;
  schoolId: SchoolId;
  thesisType: ThesisType;
  assets: ProjectAssetItem[];
  reviewResult: ReviewResponse | null;
}

const DEFAULT_MARKDOWN = `# 题目

## 摘要

请在这里输入论文摘要。

## 第一章 绪论

在这里开始撰写正文。`;


export const useDocumentStore = defineStore("document", {
  state: (): DocumentState => ({
    projectName: "paper-project",
    markdown: DEFAULT_MARKDOWN,
    schoolId: "sdfmu",
    thesisType: "thesis",
    assets: [],
    reviewResult: null,
  }),
  actions: {
    setProjectName(value: string) {
      this.projectName = value;
    },
    setMarkdown(value: string) {
      this.markdown = value;
    },
    setSchoolId(value: SchoolId) {
      this.schoolId = value;
    },
    setThesisType(value: ThesisType) {
      this.thesisType = value;
    },
    setAssets(value: ProjectAssetItem[]) {
      this.assets = value;
    },
    upsertAsset(value: ProjectAssetItem) {
      const index = this.assets.findIndex((item) => item.asset_id === value.asset_id);
      if (index === -1) {
        this.assets = [...this.assets, value];
        return;
      }

      const next = [...this.assets];
      next[index] = value;
      this.assets = next;
    },
    removeAsset(assetId: string) {
      this.assets = this.assets.filter((item) => item.asset_id !== assetId);
    },
    setReviewResult(value: ReviewResponse | null) {
      this.reviewResult = value;
    },
  },
});

