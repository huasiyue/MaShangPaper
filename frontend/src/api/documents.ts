import { http } from "./http";


export type SchoolId = "sdfmu" | "yzu" | "sdfmu_ai";
export type ThesisType = "thesis" | "design_report";

export interface AssetUploadResponse {
  asset_id: string;
  filename: string;
  url: string;
}

export interface ProjectAssetItem {
  asset_id: string;
  filename: string;
  url: string;
}

export interface ProjectImportResponse {
  filename: string;
  content: string;
  assets: ProjectAssetItem[];
}

export interface ReviewIssue {
  level: string;
  location: string;
  description: string;
  suggestion: string;
  current_value: string;
  expected_value: string;
}

export interface ReviewResponse {
  filename: string;
  school_id: SchoolId;
  thesis_type: ThesisType;
  total_issues: number;
  error_count: number;
  warning_count: number;
  info_count: number;
  issues: ReviewIssue[];
  report_text: string;
}

export async function convertMarkdown(input: {
  content: string;
  schoolId: SchoolId;
  thesisType: ThesisType;
}): Promise<Blob> {
  const formData = new FormData();
  formData.append("content", input.content);
  formData.append("school_id", input.schoolId);
  formData.append("thesis_type", input.thesisType);

  const response = await http.post("/api/documents/convert", formData, {
    responseType: "blob",
  });

  return response.data;
}

export async function reviewWordDocument(input: {
  file: File;
  schoolId: SchoolId;
  thesisType: ThesisType;
}): Promise<ReviewResponse> {
  const formData = new FormData();
  formData.append("file", input.file);
  formData.append("school_id", input.schoolId);
  formData.append("thesis_type", input.thesisType);

  const response = await http.post<ReviewResponse>("/api/documents/review", formData);
  return response.data;
}

export async function formatWordDocument(input: {
  file: File;
  schoolId: SchoolId;
  thesisType: ThesisType;
}): Promise<Blob> {
  const formData = new FormData();
  formData.append("file", input.file);
  formData.append("school_id", input.schoolId);
  formData.append("thesis_type", input.thesisType);

  const response = await http.post("/api/documents/format", formData, {
    responseType: "blob",
  });

  return response.data;
}

export async function uploadImageAsset(file: File): Promise<AssetUploadResponse> {
  const formData = new FormData();
  formData.append("file", file);

  const response = await http.post<AssetUploadResponse>("/api/assets/upload", formData);
  return response.data;
}

export async function renameImageAsset(assetId: string, filename: string): Promise<AssetUploadResponse> {
  const response = await http.patch<AssetUploadResponse>(`/api/assets/${assetId}`, {
    filename,
  });
  return response.data;
}

export async function deleteImageAsset(assetId: string): Promise<void> {
  await http.delete(`/api/assets/${assetId}`);
}

export async function exportProjectPackage(input: {
  content: string;
  schoolId: SchoolId;
  thesisType: ThesisType;
  projectName: string;
}): Promise<Blob> {
  const formData = new FormData();
  formData.append("content", input.content);
  formData.append("school_id", input.schoolId);
  formData.append("thesis_type", input.thesisType);
  formData.append("project_name", input.projectName);

  const response = await http.post("/api/documents/project/export", formData, {
    responseType: "blob",
  });
  return response.data;
}

export async function importProjectPackage(file: File): Promise<ProjectImportResponse> {
  const formData = new FormData();
  formData.append("file", file);

  const response = await http.post<ProjectImportResponse>("/api/documents/project/import", formData);
  return response.data;
}
