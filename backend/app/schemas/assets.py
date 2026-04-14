from pydantic import BaseModel, Field


class AssetUploadResponse(BaseModel):
    asset_id: str = Field(..., description="Uploaded asset id.")
    filename: str = Field(..., description="Original filename.")
    url: str = Field(..., description="Public URL for preview and reuse.")


class AssetRenameRequest(BaseModel):
    filename: str = Field(..., description="New asset filename.")
