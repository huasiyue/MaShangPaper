from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


SchoolId = Literal["sdfmu"]
ThesisTypeValue = Literal["thesis"]


class ErrorResponse(BaseModel):
    code: str = Field(..., description="Machine-readable error code.")
    message: str = Field(..., description="Human-readable error message.")


class HealthResponse(BaseModel):
    status: str
    school_support: list[SchoolId]


class ReviewIssue(BaseModel):
    level: str
    location: str
    description: str
    suggestion: str
    current_value: str = ""
    expected_value: str = ""


class ReviewResponse(BaseModel):
    filename: str
    school_id: SchoolId
    thesis_type: ThesisTypeValue
    total_issues: int
    error_count: int
    warning_count: int
    info_count: int
    issues: list[ReviewIssue]
    report_text: str


class ProjectAssetItem(BaseModel):
    asset_id: str
    filename: str
    url: str


class ProjectImportResponse(BaseModel):
    filename: str
    content: str
    assets: list[ProjectAssetItem]
