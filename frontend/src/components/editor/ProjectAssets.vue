<script setup lang="ts">
import { computed, ref } from "vue";

import type { ProjectAssetItem } from "@/api/documents";


const props = defineProps<{
  projectName: string;
  assets: ProjectAssetItem[];
  renameLoadingId: string | null;
  deleteLoadingId: string | null;
  uploadLoading: boolean;
  selectedAssetId: string | null;
}>();

const emit = defineEmits<{
  (e: "update:projectName", value: string): void;
  (e: "rename", payload: { assetId: string; filename: string }): void;
  (e: "delete", assetId: string): void;
  (e: "upload", files: File[]): void;
  (e: "insert", asset: ProjectAssetItem): void;
  (e: "select-paper"): void;
  (e: "select-asset", assetId: string): void;
}>();

const editingId = ref<string | null>(null);
const editingName = ref("");
const dragActive = ref(false);
const uploadInputRef = ref<HTMLInputElement | null>(null);

const selectedAsset = computed(() => props.assets.find((item) => item.asset_id === props.selectedAssetId) ?? null);

function startEdit(item: ProjectAssetItem) {
  editingId.value = item.asset_id;
  editingName.value = item.filename;
}

function confirmEdit() {
  if (!editingId.value || !editingName.value.trim()) {
    return;
  }

  emit("rename", {
    assetId: editingId.value,
    filename: editingName.value.trim(),
  });
  editingId.value = null;
  editingName.value = "";
}

function cancelEdit() {
  editingId.value = null;
  editingName.value = "";
}

function triggerUpload() {
  uploadInputRef.value?.click();
}

function handleFileChange(event: Event) {
  const target = event.target as HTMLInputElement;
  const files = Array.from(target.files ?? []);
  if (files.length === 0) {
    return;
  }

  emit("upload", files);
  target.value = "";
}

function handleDrop(event: DragEvent) {
  event.preventDefault();
  dragActive.value = false;
  const files = Array.from(event.dataTransfer?.files ?? []).filter((file) => file.type.startsWith("image/"));
  if (files.length > 0) {
    emit("upload", files);
  }
}
</script>

<template>
  <a-card class="assets-card panel-shell" :bordered="false">
    <template #title>项目</template>

    <a-space direction="vertical" fill size="large">
      <a-input
        :model-value="props.projectName"
        size="small"
        placeholder="项目名"
        @update:model-value="emit('update:projectName', $event)"
      />

      <div class="explorer-tree">
        <div class="tree-root">{{ props.projectName }}</div>

        <button class="tree-node tree-node--file" type="button" @click="emit('select-paper')">
          <span class="tree-node__name">paper.md</span>
        </button>

        <div class="tree-folder">
          <div class="tree-folder__header">
            <span>assets</span>
            <a-tag size="small" bordered>{{ props.assets.length }}</a-tag>
          </div>

          <div
            class="upload-dropzone"
            :class="{ 'upload-dropzone--active': dragActive }"
            @click="triggerUpload"
            @dragenter.prevent="dragActive = true"
            @dragover.prevent="dragActive = true"
            @dragleave.prevent="dragActive = false"
            @drop="handleDrop"
          >
            <span>{{ props.uploadLoading ? "上传中..." : "拖拽或点击上传" }}</span>
          </div>
          <input ref="uploadInputRef" class="hidden-input" type="file" accept="image/*" multiple @change="handleFileChange" />

          <template v-if="props.assets.length > 0">
            <div class="tree-children">
              <template v-for="item in props.assets" :key="item.asset_id">
                <div
                  class="tree-node tree-node--asset"
                  :class="{ 'tree-node--selected': props.selectedAssetId === item.asset_id }"
                  @click="emit('select-asset', item.asset_id)"
                >
                  <a-image :src="item.url" width="34" height="34" fit="cover" />

                  <div class="tree-node__body">
                    <template v-if="editingId === item.asset_id">
                      <a-input v-model="editingName" size="mini" />
                      <a-space size="mini" wrap>
                        <a-button size="mini" type="primary" :loading="props.renameLoadingId === item.asset_id" @click.stop="confirmEdit">
                          保存
                        </a-button>
                        <a-button size="mini" @click.stop="cancelEdit">取消</a-button>
                      </a-space>
                    </template>

                    <template v-else>
                      <div class="tree-node__name" :title="item.filename">{{ item.filename }}</div>
                      <a-space size="mini" wrap>
                        <a-button size="mini" @click.stop="emit('insert', item)">插入</a-button>
                        <a-button size="mini" @click.stop="startEdit(item)">改名</a-button>
                        <a-button size="mini" status="danger" :loading="props.deleteLoadingId === item.asset_id" @click.stop="emit('delete', item.asset_id)">
                          删除
                        </a-button>
                      </a-space>
                    </template>
                  </div>
                </div>
              </template>
            </div>
          </template>
        </div>
      </div>

      <div class="preview-pane">
        <template v-if="selectedAsset">
          <a-image :src="selectedAsset.url" width="100%" height="180" fit="contain" />
          <div class="preview-pane__name">{{ selectedAsset.filename }}</div>
        </template>
        <template v-else>
          <a-empty description="选择素材预览" />
        </template>
      </div>
    </a-space>
  </a-card>
</template>

<style scoped>
.assets-card {
  border-radius: 28px;
}

.explorer-tree {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.tree-root {
  font-size: 12px;
  color: var(--msp-text-muted);
  font-weight: 700;
}

.tree-folder {
  display: flex;
  flex-direction: column;
  gap: 10px;
  padding: 12px;
  border-radius: 20px;
  background: rgba(255, 255, 255, 0.6);
}

.tree-folder__header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  font-size: 13px;
  font-weight: 700;
}

.tree-children {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.tree-node {
  display: flex;
  align-items: center;
  gap: 10px;
  width: 100%;
  padding: 10px 12px;
  border: 0;
  border-radius: 18px;
  background: rgba(255, 255, 255, 0.72);
  text-align: left;
}

.tree-node--file {
  cursor: pointer;
}

.tree-node--asset {
  cursor: pointer;
}

.tree-node--selected {
  box-shadow: inset 0 0 0 1px rgba(22, 93, 255, 0.22);
  background: rgba(234, 243, 255, 0.96);
}

.tree-node__body {
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.tree-node__name {
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
  font-size: 13px;
  font-weight: 600;
}

.upload-dropzone {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 78px;
  border-radius: 18px;
  border: 1px dashed rgba(22, 93, 255, 0.24);
  background: rgba(255, 255, 255, 0.72);
  color: var(--msp-text-muted);
  font-size: 12px;
  cursor: pointer;
  transition: all 0.2s ease;
}

.upload-dropzone--active {
  border-color: rgba(22, 93, 255, 0.46);
  background: rgba(234, 243, 255, 0.9);
}

.hidden-input {
  display: none;
}

.preview-pane {
  padding: 12px;
  border-radius: 20px;
  background: rgba(255, 255, 255, 0.62);
}

.preview-pane__name {
  margin-top: 10px;
  font-size: 13px;
  font-weight: 600;
  text-align: center;
}
</style>
