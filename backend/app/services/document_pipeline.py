from __future__ import annotations

import json
import re
import shutil
from pathlib import Path
from uuid import uuid4
from zipfile import ZIP_DEFLATED, ZipFile

from fastapi import UploadFile

from app.core.settings import DEFAULT_SCHOOL_ID, DEFAULT_THESIS_TYPE, SUPPORTED_SCHOOLS, TEMP_DIR
from app.schemas.documents import ProjectAssetItem, ProjectImportResponse, ReviewResponse
from app.services.assets import ASSET_URL_PATTERN, find_asset_path, store_asset_from_path
from app.services.docx_generator import DocxGenerator, GenerationResult
from app.services.formatters.base import BaseFormatter
from app.services.formatters.schools import get_formatter_class


def cleanup_paths(paths: list[Path]) -> None:
    for path in paths:
        try:
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
            else:
                path.unlink(missing_ok=True)
        except OSError:
            continue


class DocumentPipeline:
    def __init__(self) -> None:
        self.generator = DocxGenerator()

    def _ensure_school_supported(self, school_id: str) -> None:
        if school_id not in SUPPORTED_SCHOOLS:
            raise ValueError(f"当前仅支持学校模板: {', '.join(SUPPORTED_SCHOOLS)}")

    def _get_formatter(self, school_id: str) -> BaseFormatter:
        self._ensure_school_supported(school_id)
        return get_formatter_class(school_id)()

    def _normalize_thesis_type(self, thesis_type: str) -> str:
        if thesis_type in {"design", "report"}:
            return "design_report"
        return thesis_type if thesis_type else DEFAULT_THESIS_TYPE

    def _safe_name(self, name: str) -> str:
        stem = Path(name).name or "document"
        return re.sub(r"[^A-Za-z0-9._-]+", "_", stem)

    async def persist_upload(self, upload: UploadFile) -> Path:
        safe_name = self._safe_name(upload.filename or "document.docx")
        target = TEMP_DIR / f"{uuid4().hex}_{safe_name}"
        content = await upload.read()
        target.write_bytes(content)
        await upload.close()
        return target

    def convert_markdown(
        self,
        content: str,
        school_id: str = DEFAULT_SCHOOL_ID,
        thesis_type: str = DEFAULT_THESIS_TYPE,
    ) -> GenerationResult:
        self._ensure_school_supported(school_id)
        normalized_type = self._normalize_thesis_type(thesis_type)
        token = uuid4().hex
        markdown_path = TEMP_DIR / f"{token}_{normalized_type}.md"
        output_path = TEMP_DIR / f"{token}_{normalized_type}.docx"
        return self.generator.generate(content=content, markdown_path=markdown_path, output_path=output_path)

    def review_document(
        self,
        input_path: Path,
        school_id: str = DEFAULT_SCHOOL_ID,
        thesis_type: str = DEFAULT_THESIS_TYPE,
    ) -> ReviewResponse:
        normalized_type = self._normalize_thesis_type(thesis_type)
        formatter = self._get_formatter(school_id)
        return formatter.review_document(input_path=input_path, thesis_type=normalized_type)

    def format_document(
        self,
        input_path: Path,
        school_id: str = DEFAULT_SCHOOL_ID,
        thesis_type: str = DEFAULT_THESIS_TYPE,
    ) -> tuple[Path, Path, ReviewResponse]:
        normalized_type = self._normalize_thesis_type(thesis_type)
        formatter = self._get_formatter(school_id)
        formatted_path = TEMP_DIR / f"{uuid4().hex}_{input_path.stem}_formatted.docx"
        saved_path, review = formatter.format_document(
            input_path=input_path,
            output_path=formatted_path,
            thesis_type=normalized_type,
        )

        archive_path = TEMP_DIR / f"{uuid4().hex}_{input_path.stem}_formatted.zip"
        with ZipFile(archive_path, mode="w", compression=ZIP_DEFLATED) as archive:
            archive.write(saved_path, arcname=f"{input_path.stem}_formatted.docx")
            archive.writestr("review.json", json.dumps(review.model_dump(), ensure_ascii=False, indent=2))
            archive.writestr("review.txt", review.report_text)

        return saved_path, archive_path, review

    def export_project_package(self, content: str, school_id: str, thesis_type: str, project_name: str = "paper-project") -> Path:
        self._ensure_school_supported(school_id)
        safe_project_name = re.sub(r"[^A-Za-z0-9._-]+", "_", project_name).strip("._") or "paper-project"
        archive_path = TEMP_DIR / f"{uuid4().hex}_{safe_project_name}.zip"

        replacements: dict[str, str] = {}
        used_names: set[str] = set()

        def unique_name(original: str) -> str:
            base_path = Path(original)
            stem = base_path.stem or "asset"
            suffix = base_path.suffix
            candidate = f"{stem}{suffix}"
            index = 1
            while candidate in used_names:
                candidate = f"{stem}_{index}{suffix}"
                index += 1
            used_names.add(candidate)
            return candidate

        def replace_url(match: re.Match[str]) -> str:
            asset_id = match.group("asset_id") or match.group("asset_id_rel")
            if not asset_id:
                return match.group(0)

            try:
                asset_path = find_asset_path(asset_id)
            except FileNotFoundError:
                return match.group(0)

            if asset_id not in replacements:
                replacements[asset_id] = unique_name(asset_path.name)
            return f"assets/{replacements[asset_id]}"

        package_content = ASSET_URL_PATTERN.sub(replace_url, content)

        with ZipFile(archive_path, mode="w", compression=ZIP_DEFLATED) as archive:
            archive.writestr("paper.md", package_content)
            for asset_id, packaged_name in replacements.items():
                archive.write(find_asset_path(asset_id), arcname=f"assets/{packaged_name}")

        return archive_path

    def import_project_package(self, zip_path: Path) -> tuple[ProjectImportResponse, Path]:
        extract_dir = TEMP_DIR / f"{uuid4().hex}_project"
        extract_dir.mkdir(parents=True, exist_ok=True)

        with ZipFile(zip_path) as archive:
            archive.extractall(extract_dir)

        markdown_files = sorted(path for path in extract_dir.rglob("*.md") if path.is_file())
        if not markdown_files:
            raise ValueError("项目包中未找到 Markdown 文件。")

        markdown_path = markdown_files[0]
        content = markdown_path.read_text(encoding="utf-8")
        assets: list[ProjectAssetItem] = []
        asset_by_name: dict[str, ProjectAssetItem] = {}

        for asset_path in sorted(
            path for path in extract_dir.rglob("*")
            if path.is_file() and path.suffix.lower() not in {".md"}
        ):
            stored = store_asset_from_path(asset_path)
            item = ProjectAssetItem(
                asset_id=stored.asset_id,
                filename=stored.filename,
                url=f"/api/assets/{stored.asset_id}",
            )
            assets.append(item)
            asset_by_name[asset_path.name] = item

        def replace_local_image(match: re.Match[str]) -> str:
            alt_text = match.group(1)
            target = match.group(2).strip()
            if target.startswith("http://") or target.startswith("https://") or target.startswith("/api/assets/"):
                return match.group(0)

            item = asset_by_name.get(Path(target).name)
            if not item:
                return match.group(0)
            return f"![{alt_text}]({item.url})"

        normalized_content = re.sub(r"!\[(.*?)\]\((.*?)\)", replace_local_image, content)

        response = ProjectImportResponse(
            filename=zip_path.name,
            content=normalized_content,
            assets=assets,
        )
        return response, extract_dir
