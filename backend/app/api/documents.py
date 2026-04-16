from __future__ import annotations

from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask

from app.core.settings import DEFAULT_SCHOOL_ID, DEFAULT_THESIS_TYPE, SUPPORTED_SCHOOLS
from app.schemas.documents import ErrorResponse, HealthResponse, ProjectAssetItem, ProjectImportResponse, ReviewResponse
from app.services.document_pipeline import DocumentPipeline, cleanup_paths


router = APIRouter(prefix="/api/documents", tags=["documents"])
pipeline = DocumentPipeline()


def error_payload(code: str, message: str) -> dict[str, str]:
    return {"code": code, "message": message}


def ensure_word_file(path: Path) -> None:
    if path.suffix.lower() not in {".doc", ".docx"}:
        raise HTTPException(
            status_code=400,
            detail=error_payload("invalid_file_type", "仅支持 .doc 或 .docx 文件。"),
        )


@router.get(
    "/health",
    response_model=HealthResponse,
    responses={500: {"model": ErrorResponse}},
)
def health() -> HealthResponse:
    return HealthResponse(status="ok", school_support=list(SUPPORTED_SCHOOLS))


@router.post(
    "/convert",
    responses={
        400: {"model": ErrorResponse},
        500: {"model": ErrorResponse},
    },
)
async def convert_document(
    content: str = Form(...),
    school_id: str = Form(DEFAULT_SCHOOL_ID),
    thesis_type: str = Form(DEFAULT_THESIS_TYPE),
):
    if not content.strip():
        raise HTTPException(
            status_code=400,
            detail=error_payload("empty_content", "Markdown 内容不能为空。"),
        )

    try:
        result = pipeline.convert_markdown(content=content, school_id=school_id, thesis_type=thesis_type)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=error_payload("unsupported_school", str(exc))) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=error_payload("internal_error", str(exc))) from exc

    filename = f"{school_id}_{thesis_type}_draft.docx"
    return FileResponse(
        path=result.output_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        background=BackgroundTask(cleanup_paths, [result.markdown_path, result.output_path]),
    )


@router.post(
    "/project/export",
    responses={
        400: {"model": ErrorResponse},
        500: {"model": ErrorResponse},
    },
)
async def export_project_package(
    content: str = Form(...),
    school_id: str = Form(DEFAULT_SCHOOL_ID),
    thesis_type: str = Form(DEFAULT_THESIS_TYPE),
    project_name: str = Form("paper-project"),
):
    if not content.strip():
        raise HTTPException(
            status_code=400,
            detail=error_payload("empty_content", "Markdown 内容不能为空。"),
        )

    try:
        archive_path = pipeline.export_project_package(
            content=content,
            school_id=school_id,
            thesis_type=thesis_type,
            project_name=project_name,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=error_payload("bad_request", str(exc))) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=error_payload("internal_error", str(exc))) from exc

    filename = f"{project_name or 'paper-project'}.zip"
    return FileResponse(
        path=archive_path,
        filename=filename,
        media_type="application/zip",
        background=BackgroundTask(cleanup_paths, [archive_path]),
    )


@router.post(
    "/project/import",
    response_model=ProjectImportResponse,
    responses={
        400: {"model": ErrorResponse},
        500: {"model": ErrorResponse},
    },
)
async def import_project_package(request: Request, file: UploadFile = File(...)) -> ProjectImportResponse:
    stored_path = await pipeline.persist_upload(file)

    if stored_path.suffix.lower() != ".zip":
        cleanup_paths([stored_path])
        raise HTTPException(
            status_code=400,
            detail=error_payload("invalid_file_type", "仅支持 .zip 项目包。"),
        )

    extract_dir: Path | None = None
    try:
        project, extract_dir = pipeline.import_project_package(stored_path)
        base_url = str(request.base_url).rstrip("/")
        normalized_assets = [
            ProjectAssetItem(
                asset_id=asset.asset_id,
                filename=asset.filename,
                url=f"{base_url}/api/assets/{asset.asset_id}",
            )
            for asset in project.assets
        ]
        normalized_content = project.content
        for asset in normalized_assets:
            normalized_content = normalized_content.replace(f"/api/assets/{asset.asset_id}", asset.url)
        cleanup_paths([stored_path, extract_dir])
        return ProjectImportResponse(
            filename=project.filename,
            content=normalized_content,
            assets=normalized_assets,
        )
    except ValueError as exc:
        cleanup_paths([stored_path, extract_dir])
        raise HTTPException(status_code=400, detail=error_payload("bad_request", str(exc))) from exc
    except Exception as exc:
        cleanup_paths([stored_path, extract_dir])
        raise HTTPException(status_code=500, detail=error_payload("internal_error", str(exc))) from exc


@router.post(
    "/review",
    response_model=ReviewResponse,
    responses={
        400: {"model": ErrorResponse},
        500: {"model": ErrorResponse},
    },
)
async def review_document(
    file: UploadFile = File(...),
    school_id: str = Form(DEFAULT_SCHOOL_ID),
    thesis_type: str = Form(DEFAULT_THESIS_TYPE),
) -> ReviewResponse:
    stored_path = await pipeline.persist_upload(file)

    try:
        ensure_word_file(stored_path)
        return pipeline.review_document(
            input_path=stored_path,
            school_id=school_id,
            thesis_type=thesis_type,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=error_payload("bad_request", str(exc))) from exc
    finally:
        cleanup_paths([stored_path])


@router.post(
    "/format",
    responses={
        400: {"model": ErrorResponse},
        500: {"model": ErrorResponse},
    },
)
async def format_document(
    file: UploadFile = File(...),
    school_id: str = Form(DEFAULT_SCHOOL_ID),
    thesis_type: str = Form(DEFAULT_THESIS_TYPE),
):
    stored_path = await pipeline.persist_upload(file)

    formatted_path: Path | None = None
    archive_path: Path | None = None
    try:
        ensure_word_file(stored_path)
        formatted_path, archive_path, _review = pipeline.format_document(
            input_path=stored_path,
            school_id=school_id,
            thesis_type=thesis_type,
        )
    except ValueError as exc:
        cleanup_paths([stored_path, formatted_path, archive_path])
        raise HTTPException(status_code=400, detail=error_payload("bad_request", str(exc))) from exc
    except Exception as exc:
        cleanup_paths([stored_path, formatted_path, archive_path])
        raise HTTPException(status_code=500, detail=error_payload("internal_error", str(exc))) from exc

    filename = f"{stored_path.stem}_formatted_bundle.zip"
    return FileResponse(
        path=archive_path,
        filename=filename,
        media_type="application/zip",
        background=BackgroundTask(cleanup_paths, [stored_path, formatted_path, archive_path]),
    )
