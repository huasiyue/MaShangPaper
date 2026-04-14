from __future__ import annotations

from fastapi import APIRouter, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse

from app.schemas.assets import AssetRenameRequest, AssetUploadResponse
from app.schemas.documents import ErrorResponse
from app.services.assets import delete_asset, find_asset_path, rename_asset, save_asset


router = APIRouter(prefix="/api/assets", tags=["assets"])


@router.post(
    "/upload",
    response_model=AssetUploadResponse,
    responses={400: {"model": ErrorResponse}, 500: {"model": ErrorResponse}},
)
async def upload_asset(request: Request, file: UploadFile = File(...)) -> AssetUploadResponse:
    if not file.filename:
        raise HTTPException(status_code=400, detail={"code": "missing_file", "message": "请先选择图片文件。"})

    stored = await save_asset(file)
    url = str(request.base_url).rstrip("/") + f"/api/assets/{stored.asset_id}"
    return AssetUploadResponse(asset_id=stored.asset_id, filename=stored.filename, url=url)


@router.get(
    "/{asset_id}",
    responses={404: {"model": ErrorResponse}},
)
def get_asset(asset_id: str):
    try:
        path = find_asset_path(asset_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail={"code": "asset_not_found", "message": "图片资源不存在。"}) from exc

    return FileResponse(path=path)


@router.patch(
    "/{asset_id}",
    response_model=AssetUploadResponse,
    responses={404: {"model": ErrorResponse}},
)
def update_asset(request: Request, asset_id: str, payload: AssetRenameRequest) -> AssetUploadResponse:
    try:
        stored = rename_asset(asset_id, payload.filename)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail={"code": "asset_not_found", "message": "图片资源不存在。"}) from exc

    url = str(request.base_url).rstrip("/") + f"/api/assets/{stored.asset_id}"
    return AssetUploadResponse(asset_id=stored.asset_id, filename=stored.filename, url=url)


@router.delete(
    "/{asset_id}",
    responses={404: {"model": ErrorResponse}},
)
def remove_asset(asset_id: str):
    try:
        delete_asset(asset_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail={"code": "asset_not_found", "message": "图片资源不存在。"}) from exc

    return {"status": "ok"}
