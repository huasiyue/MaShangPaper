<script setup lang="ts">
import { AxiosError } from "axios";
import { marked } from "marked";
import { computed, nextTick, ref } from "vue";
import { Message } from "@arco-design/web-vue";
import "katex/dist/katex.min.css";

import {
  convertMarkdown,
  deleteImageAsset,
  exportProjectPackage,
  formatWordDocument,
  importProjectPackage,
  renameImageAsset,
  reviewWordDocument,
  uploadImageAsset,
  type ProjectAssetItem,
  type SchoolId,
  type ThesisType,
} from "@/api/documents";
import ExportPanel from "@/components/editor/ExportPanel.vue";
import ProjectAssets from "@/components/editor/ProjectAssets.vue";
import SchoolSelector from "@/components/editor/SchoolSelector.vue";
import ReviewResultTable from "@/components/review/ReviewResultTable.vue";
import { useDocumentStore } from "@/stores/document";
import { extractMath, restoreMath } from "@/utils/markdown-math";


marked.setOptions({
  gfm: true,
  breaks: true,
});

const documentStore = useDocumentStore();

const convertLoading = ref(false);
const reviewLoading = ref(false);
const formatLoading = ref(false);
const exportProjectLoading = ref(false);
const importProjectLoading = ref(false);
const imageLoading = ref(false);
const renameAssetLoadingId = ref<string | null>(null);
const deleteAssetLoadingId = ref<string | null>(null);
const selectedAssetId = ref<string | null>(null);
const selectedWordFile = ref<File | null>(null);
const markdownEditorRef = ref<any>(null);
const imageInputRef = ref<HTMLInputElement | null>(null);
const projectInputRef = ref<HTMLInputElement | null>(null);
const projectFolderInputRef = ref<HTMLInputElement | null>(null);

const markdownEmpty = computed(() => documentStore.markdown.trim().length === 0);
const wordMissing = computed(() => !selectedWordFile.value);

function updateSchoolId(value: SchoolId) {
  documentStore.setSchoolId(value);
}

function updateThesisType(value: ThesisType) {
  documentStore.setThesisType(value);
}

function updateProjectName(value: string) {
  documentStore.setProjectName(value.trim() ? value : "paper-project");
}

function readApiError(error: unknown): string {
  if (error instanceof AxiosError) {
    const data = error.response?.data;
    if (data instanceof Blob) {
      return "请求失败，请检查后端日志。";
    }
    if (typeof data?.message === "string") {
      return data.message;
    }
  }
  return error instanceof Error ? error.message : "请求失败，请稍后重试。";
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  window.URL.revokeObjectURL(url);
}

function getTextareaElement(): HTMLTextAreaElement | null {
  const root = markdownEditorRef.value?.$el as HTMLElement | undefined;
  return root?.querySelector("textarea") ?? null;
}

function insertSnippet(snippet: string) {
  const current = documentStore.markdown;
  const textarea = getTextareaElement();

  if (!textarea) {
    const spacer = current.trim() ? "\n\n" : "";
    documentStore.setMarkdown(`${current}${spacer}${snippet}`);
    return;
  }

  const start = textarea.selectionStart ?? current.length;
  const end = textarea.selectionEnd ?? start;
  const next = `${current.slice(0, start)}${snippet}${current.slice(end)}`;
  documentStore.setMarkdown(next);

  nextTick(() => {
    const target = getTextareaElement();
    if (!target) {
      return;
    }

    const position = start + snippet.length;
    target.focus();
    target.setSelectionRange(position, position);
  });
}

function renderMarkdownWithImageMeta(markdown: string): string {
  // Step 1: extract math ($...$ / $$...$$) into safe placeholders
  const { text: mathStripped, blocks } = extractMath(markdown);

  // Step 2: handle image metadata syntax
  const normalized = mathStripped.replace(
    /!\[(.*?)\]\((.*?)\)/g,
    (_match, rawAlt: string, rawUrl: string) => {
      const parts = rawAlt
        .split("|")
        .map((part) => part.trim())
        .filter(Boolean);

      const caption = parts[0] ?? "";
      let width = "12cm";
      let align = "center";

      for (const option of parts.slice(1)) {
        const [key, value] = option.split("=").map((item) => item?.trim());
        if (!key || !value) {
          continue;
        }

        if (key === "width" || key === "w") {
          width = value;
        }
        if (key === "align" || key === "a") {
          align = value;
        }
      }

      const justify = align === "left" ? "flex-start" : align === "right" ? "flex-end" : "center";
      const captionHtml = caption ? `<figcaption>${caption}</figcaption>` : "";
      return `<figure class="msp-figure" style="justify-content:${justify};"><div class="msp-figure__inner" style="width:${width};max-width:100%;"><img src="${rawUrl}" alt="${caption}" />${captionHtml}</div></figure>`;
    },
  );

  // Step 3: parse with marked
  const html = marked.parse(normalized) as string;

  // Step 4: restore math blocks with KaTeX-rendered HTML
  return restoreMath(html, blocks);
}

const previewHtml = computed(() => renderMarkdownWithImageMeta(documentStore.markdown));

function insertTableTemplate() {
  insertSnippet("\n| 列1 | 列2 | 列3 |\n| --- | --- | --- |\n| 内容 | 内容 | 内容 |\n| 内容 | 内容 | 内容 |\n");
}

function triggerImagePicker() {
  imageInputRef.value?.click();
}

function triggerProjectPicker() {
  projectInputRef.value?.click();
}

function triggerProjectFolderPicker() {
  projectFolderInputRef.value?.click();
}

function normalizeRelativePath(path: string): string {
  const parts = path.replace(/\\/g, "/").split("/");
  const normalized: string[] = [];

  for (const part of parts) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      normalized.pop();
      continue;
    }
    normalized.push(part);
  }

  return normalized.join("/");
}

function getRootRelativePath(file: File): string {
  const rawPath = file.webkitRelativePath || file.name;
  const segments = rawPath.replace(/\\/g, "/").split("/");
  if (segments.length > 1) {
    return normalizeRelativePath(segments.slice(1).join("/"));
  }
  return normalizeRelativePath(rawPath);
}

function getDirectoryPath(path: string): string {
  const normalized = normalizeRelativePath(path);
  const segments = normalized ? normalized.split("/") : [];
  segments.pop();
  return segments.join("/");
}

function resolveMarkdownAssetPath(markdownPath: string, targetPath: string): string {
  const baseDir = getDirectoryPath(markdownPath);
  const baseSegments = baseDir ? baseDir.split("/") : [];
  const targetSegments = targetPath.replace(/\\/g, "/").split("/");
  const resolved = [...baseSegments];

  for (const segment of targetSegments) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }

  return normalizeRelativePath(resolved.join("/"));
}

function isImageFile(file: File): boolean {
  return file.type.startsWith("image/") || /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(file.name);
}

function rewriteMarkdownImageTargets(
  markdown: string,
  markdownRelativePath: string,
  assetMap: Map<string, ProjectAssetItem>,
): string {
  return markdown.replace(/!\[(.*?)\]\((.*?)\)/g, (match, altText: string, rawTarget: string) => {
    const cleanedTarget = rawTarget.trim().replace(/^<|>$/g, "").replace(/^["']|["']$/g, "");
    if (
      /^(?:https?:\/\/|data:|\/api\/assets\/)/i.test(cleanedTarget) ||
      cleanedTarget.startsWith("#")
    ) {
      return match;
    }

    const resolvedPath = resolveMarkdownAssetPath(markdownRelativePath, cleanedTarget);
    const asset = assetMap.get(resolvedPath);
    if (!asset) {
      return match;
    }
    return `![${altText}](${asset.url})`;
  });
}

function handleSelectPaper() {
  selectedAssetId.value = null;
  nextTick(() => {
    getTextareaElement()?.focus();
  });
}

async function handleConvert() {
  if (markdownEmpty.value) {
    Message.warning("请先输入 Markdown 内容。");
    return;
  }

  convertLoading.value = true;
  try {
    const blob = await convertMarkdown({
      content: documentStore.markdown,
      schoolId: documentStore.schoolId,
      thesisType: documentStore.thesisType,
    });
    downloadBlob(blob, `${documentStore.schoolId}_${documentStore.thesisType}_draft.docx`);
    Message.success("已导出。");
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    convertLoading.value = false;
  }
}

async function handleReview() {
  if (!selectedWordFile.value) {
    Message.warning("请先选择 Word 文件。");
    return;
  }

  reviewLoading.value = true;
  try {
    const review = await reviewWordDocument({
      file: selectedWordFile.value,
      schoolId: documentStore.schoolId,
      thesisType: documentStore.thesisType,
    });
    documentStore.setReviewResult(review);
    Message.success("已审查。");
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    reviewLoading.value = false;
  }
}

async function handleFormat() {
  if (!selectedWordFile.value) {
    Message.warning("请先选择 Word 文件。");
    return;
  }

  formatLoading.value = true;
  try {
    const blob = await formatWordDocument({
      file: selectedWordFile.value,
      schoolId: documentStore.schoolId,
      thesisType: documentStore.thesisType,
    });
    downloadBlob(blob, `${documentStore.schoolId}_${documentStore.thesisType}_formatted_bundle.zip`);
    Message.success("已格式化。");
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    formatLoading.value = false;
  }
}

async function handleExportProject() {
  if (markdownEmpty.value) {
    Message.warning("请先输入 Markdown 内容。");
    return;
  }

  exportProjectLoading.value = true;
  try {
    const blob = await exportProjectPackage({
      content: documentStore.markdown,
      schoolId: documentStore.schoolId,
      thesisType: documentStore.thesisType,
      projectName: documentStore.projectName,
    });
    downloadBlob(blob, `${documentStore.projectName || "paper-project"}.zip`);
    Message.success("项目包已导出。");
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    exportProjectLoading.value = false;
  }
}

function handleWordChange(event: Event) {
  const target = event.target as HTMLInputElement;
  selectedWordFile.value = target.files?.[0] ?? null;
}

async function handleImageChange(event: Event) {
  const target = event.target as HTMLInputElement;
  const files = Array.from(target.files ?? []);
  if (files.length === 0) {
    return;
  }

  imageLoading.value = true;
  try {
    const [firstFile] = files;
    for (const file of files) {
      const asset = await uploadImageAsset(file);
      documentStore.upsertAsset(asset);
      selectedAssetId.value = asset.asset_id;
    }

    if (firstFile) {
      const latest = documentStore.assets[documentStore.assets.length - 1];
      const caption = firstFile.name.replace(/\.[^.]+$/, "") || "图片";
      if (latest) {
        insertSnippet(`\n![${caption}|width=12cm|align=center](${latest.url})\n`);
      }
    }

    Message.success("素材已上传。");
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    imageLoading.value = false;
    target.value = "";
  }
}

async function handleProjectChange(event: Event) {
  const target = event.target as HTMLInputElement;
  const file = target.files?.[0];
  if (!file) {
    return;
  }

  importProjectLoading.value = true;
  try {
    const project = await importProjectPackage(file);
    documentStore.setProjectName(project.filename.replace(/\.zip$/i, "") || "paper-project");
    documentStore.setMarkdown(project.content);
    documentStore.setAssets(project.assets);
    selectedAssetId.value = project.assets[0]?.asset_id ?? null;
    Message.success("项目包已导入。");
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    importProjectLoading.value = false;
    target.value = "";
  }
}

async function handleProjectFolderChange(event: Event) {
  const target = event.target as HTMLInputElement;
  const files = Array.from(target.files ?? []);
  if (files.length === 0) {
    return;
  }

  importProjectLoading.value = true;
  try {
    const sortedFiles = [...files].sort((left, right) =>
      getRootRelativePath(left).localeCompare(getRootRelativePath(right), "zh-CN"),
    );
    const markdownFiles = sortedFiles.filter((file) => file.name.toLowerCase().endsWith(".md"));
    if (markdownFiles.length === 0) {
      Message.warning("所选文件夹中未找到 Markdown 文件。");
      return;
    }

    const markdownFile = markdownFiles[0];
    const rootName = (markdownFile.webkitRelativePath || markdownFile.name).split("/")[0] || "paper-project";
    const markdownRelativePath = getRootRelativePath(markdownFile);
    const imageFiles = sortedFiles.filter((file) => isImageFile(file));
    const uploadedAssets: ProjectAssetItem[] = [];
    const assetMap = new Map<string, ProjectAssetItem>();

    for (const file of imageFiles) {
      const asset = await uploadImageAsset(file);
      uploadedAssets.push(asset);
      assetMap.set(getRootRelativePath(file), asset);
    }

    let markdownContent = await markdownFile.text();
    markdownContent = rewriteMarkdownImageTargets(markdownContent, markdownRelativePath, assetMap);

    documentStore.setProjectName(rootName.replace(/\.[^.]+$/, "") || "paper-project");
    documentStore.setMarkdown(markdownContent);
    documentStore.setAssets(uploadedAssets);
    selectedAssetId.value = uploadedAssets[0]?.asset_id ?? null;

    if (markdownFiles.length > 1) {
      Message.success(`文件夹已导入，已使用 ${markdownFile.name}，共导入 ${uploadedAssets.length} 张图片。`);
    } else {
      Message.success(`文件夹已导入，共导入 ${uploadedAssets.length} 张图片。`);
    }
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    importProjectLoading.value = false;
    target.value = "";
  }
}

function removeAssetMarkdownReferences(asset: ProjectAssetItem) {
  const escapedUrl = asset.url.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const next = documentStore.markdown
    .replace(new RegExp(`!?\\[[^\\]]*\\]\\(${escapedUrl}\\)\\n?`, "g"), "")
    .replace(/\n{3,}/g, "\n\n");
  documentStore.setMarkdown(next.trimEnd());
}

async function handleAssetRename(payload: { assetId: string; filename: string }) {
  renameAssetLoadingId.value = payload.assetId;
  try {
    const asset = await renameImageAsset(payload.assetId, payload.filename);
    documentStore.upsertAsset(asset);
    Message.success("素材已改名。");
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    renameAssetLoadingId.value = null;
  }
}

async function handleAssetDelete(assetId: string) {
  const asset = documentStore.assets.find((item) => item.asset_id === assetId);
  if (!asset) {
    return;
  }

  deleteAssetLoadingId.value = assetId;
  try {
    await deleteImageAsset(assetId);
    documentStore.removeAsset(assetId);
    removeAssetMarkdownReferences(asset);
    if (selectedAssetId.value === assetId) {
      selectedAssetId.value = null;
    }
    Message.success("素材已删除。");
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    deleteAssetLoadingId.value = null;
  }
}

async function handleExplorerUpload(files: File[]) {
  imageLoading.value = true;
  try {
    for (const file of files) {
      const asset = await uploadImageAsset(file);
      documentStore.upsertAsset(asset);
      selectedAssetId.value = asset.asset_id;
    }
    Message.success("素材已上传。");
  } catch (error) {
    Message.error(readApiError(error));
  } finally {
    imageLoading.value = false;
  }
}

function handleAssetSelect(assetId: string) {
  selectedAssetId.value = assetId;
}

function handleAssetInsert(asset: ProjectAssetItem) {
  const caption = asset.filename.replace(/\.[^.]+$/, "") || "图片";
  selectedAssetId.value = asset.asset_id;
  insertSnippet(`\n![${caption}|width=12cm|align=center](${asset.url})\n`);
  Message.success("已插入到 Markdown。");
}
</script>

<template>
  <div class="editor-page">
    <div class="toolbar panel-shell">
      <div class="toolbar__brand">
        <span class="toolbar__mark"></span>
        <span>MaShangPaper</span>
      </div>

      <div class="toolbar__controls">
        <a-input
          :model-value="documentStore.projectName"
          size="small"
          class="project-name-input"
          placeholder="项目名"
          @update:model-value="updateProjectName"
        />

        <SchoolSelector
          :school-id="documentStore.schoolId"
          :thesis-type="documentStore.thesisType"
          @update:school-id="updateSchoolId"
          @update:thesis-type="updateThesisType"
        />

        <a-button size="small" :loading="importProjectLoading" @click="triggerProjectPicker">
          导入压缩包
        </a-button>
        <input ref="projectInputRef" class="file-input" type="file" accept=".zip" @change="handleProjectChange" />

        <a-button size="small" :loading="importProjectLoading" @click="triggerProjectFolderPicker">
          导入文件夹
        </a-button>
        <input
          ref="projectFolderInputRef"
          class="file-input"
          type="file"
          multiple
          webkitdirectory
          directory
          @change="handleProjectFolderChange"
        />

        <label class="file-chip" for="word-file-input">
          <span>{{ selectedWordFile?.name ?? "选择 Word(.doc/.docx)" }}</span>
        </label>
        <input id="word-file-input" class="file-input" type="file" accept=".doc,.docx" @change="handleWordChange" />

        <ExportPanel
          :export-project-loading="exportProjectLoading"
          :convert-loading="convertLoading"
          :review-loading="reviewLoading"
          :format-loading="formatLoading"
          :markdown-disabled="markdownEmpty"
          :word-disabled="wordMissing"
          @export-project="handleExportProject"
          @convert="handleConvert"
          @review="handleReview"
          @format="handleFormat"
        />
      </div>
    </div>

    <div class="workbench">
      <a-card class="panel-shell editor-card" :bordered="false">
        <template #title>编辑</template>
        <template #extra>
          <a-space size="small">
            <a-button size="mini" :loading="imageLoading" @click="triggerImagePicker">
              插图
            </a-button>
            <a-button size="mini" @click="insertTableTemplate">
              表格
            </a-button>
          </a-space>
        </template>
        <input ref="imageInputRef" class="file-input" type="file" accept="image/*" @change="handleImageChange" />
        <a-textarea
          ref="markdownEditorRef"
          :model-value="documentStore.markdown"
          :auto-size="{ minRows: 26, maxRows: 34 }"
          placeholder="输入 Markdown"
          @update:model-value="documentStore.setMarkdown($event)"
        />
      </a-card>

      <a-card class="panel-shell preview-card" :bordered="false">
        <template #title>预览</template>
        <div class="markdown-preview" v-html="previewHtml" />
      </a-card>
    </div>

    <div class="result-grid">
      <ProjectAssets
        :project-name="documentStore.projectName"
        :assets="documentStore.assets"
        :rename-loading-id="renameAssetLoadingId"
        :delete-loading-id="deleteAssetLoadingId"
        :upload-loading="imageLoading"
        :selected-asset-id="selectedAssetId"
        @update:project-name="updateProjectName"
        @rename="handleAssetRename"
        @delete="handleAssetDelete"
        @upload="handleExplorerUpload"
        @insert="handleAssetInsert"
        @select-paper="handleSelectPaper"
        @select-asset="handleAssetSelect"
      />

      <ReviewResultTable :review="documentStore.reviewResult" />
    </div>
  </div>
</template>

<style scoped>
.editor-page {
  max-width: 1380px;
  margin: 0 auto;
  padding: 24px 18px 48px;
  display: flex;
  flex-direction: column;
  gap: 18px;
}

.panel-shell {
  border-radius: 28px;
  background: linear-gradient(180deg, rgba(255, 255, 255, 0.98), rgba(248, 251, 255, 0.92));
  box-shadow: 0 18px 54px rgba(15, 23, 42, 0.07);
  border: 1px solid rgba(255, 255, 255, 0.7);
}

.toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 20px;
  padding: 18px 22px;
  flex-wrap: wrap;
}

.toolbar__brand {
  display: flex;
  align-items: center;
  gap: 12px;
  font-size: 30px;
  font-weight: 800;
  letter-spacing: -0.03em;
}

.toolbar__mark {
  width: 12px;
  height: 12px;
  border-radius: 999px;
  background: linear-gradient(135deg, #165dff, #0fc6c2);
  box-shadow: 0 0 0 8px rgba(22, 93, 255, 0.08);
}

.toolbar__controls {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 12px;
  flex-wrap: wrap;
}

.project-name-input {
  width: 180px;
}

.file-input {
  display: none;
}

.file-chip {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 32px;
  padding: 0 14px;
  border-radius: 999px;
  border: 1px dashed rgba(22, 93, 255, 0.26);
  background: rgba(255, 255, 255, 0.74);
  color: var(--msp-text-muted);
  font-size: 12px;
  cursor: pointer;
  transition: all 0.2s ease;
}

.file-chip:hover {
  border-color: rgba(22, 93, 255, 0.45);
  color: var(--msp-text);
}

.workbench {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 18px;
}

.result-grid {
  display: grid;
  grid-template-columns: 320px minmax(0, 1fr);
  gap: 18px;
}

.editor-card,
.preview-card {
  min-height: 640px;
}

.markdown-preview {
  min-height: 610px;
  max-height: 720px;
  overflow: auto;
  padding-right: 6px;
  line-height: 1.82;
}

.markdown-preview :deep(h1),
.markdown-preview :deep(h2),
.markdown-preview :deep(h3) {
  margin-top: 1.05em;
  margin-bottom: 0.42em;
}

.markdown-preview :deep(.msp-figure) {
  display: flex;
  margin: 18px 0;
}

.markdown-preview :deep(.msp-figure__inner) {
  max-width: 100%;
}

.markdown-preview :deep(.msp-figure img) {
  display: block;
  width: 100%;
  max-width: 100%;
  border-radius: 16px;
}

.markdown-preview :deep(.msp-figure figcaption) {
  margin-top: 8px;
  text-align: center;
  font-size: 12px;
  color: var(--msp-text-muted);
}

.markdown-preview :deep(pre) {
  overflow: auto;
  padding: 16px;
  border-radius: 18px;
  background: rgba(15, 23, 42, 0.92);
  color: #f8fafc;
}

.markdown-preview :deep(table) {
  width: 100%;
  border-collapse: collapse;
}

.markdown-preview :deep(.katex-display) {
  margin: 16px 0;
  overflow-x: auto;
  overflow-y: hidden;
  padding: 4px 0;
}

.markdown-preview :deep(.msp-math-error) {
  color: var(--msp-text-muted);
  background: rgba(255, 0, 0, 0.06);
  padding: 2px 6px;
  border-radius: 4px;
  font-size: 12px;
}

.markdown-preview :deep(th),
.markdown-preview :deep(td) {
  padding: 10px 12px;
  border-bottom: 1px solid rgba(15, 23, 42, 0.08);
}

@media (max-width: 1120px) {
  .toolbar__brand {
    font-size: 24px;
  }

  .workbench {
    grid-template-columns: 1fr;
  }

  .result-grid {
    grid-template-columns: 1fr;
  }

  .editor-card,
  .preview-card {
    min-height: auto;
  }
}
</style>
