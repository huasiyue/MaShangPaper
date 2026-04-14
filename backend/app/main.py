from __future__ import annotations

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.api.assets import router as assets_router
from app.api.documents import router as documents_router
from app.core.settings import SUPPORTED_SCHOOLS
from app.schemas.documents import ErrorResponse, HealthResponse


def create_app() -> FastAPI:
    app = FastAPI(
        title="MaShangPaper API",
        version="0.1.0",
        summary="Markdown 论文转 Word 与格式审查服务",
    )

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["http://127.0.0.1:5173", "http://localhost:5173"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @app.exception_handler(HTTPException)
    async def handle_http_exception(_request: Request, exc: HTTPException) -> JSONResponse:
        detail = exc.detail
        if not isinstance(detail, dict):
            detail = ErrorResponse(code="http_error", message=str(detail)).model_dump()
        return JSONResponse(status_code=exc.status_code, content=detail)

    @app.exception_handler(Exception)
    async def handle_unexpected_exception(_request: Request, _exc: Exception) -> JSONResponse:
        detail = ErrorResponse(code="internal_error", message="服务器内部错误，请稍后重试。")
        return JSONResponse(status_code=500, content=detail.model_dump())

    @app.get("/", response_model=HealthResponse, tags=["system"])
    async def root() -> HealthResponse:
        return HealthResponse(status="ok", school_support=list(SUPPORTED_SCHOOLS))

    app.include_router(documents_router)
    app.include_router(assets_router)
    return app


app = create_app()
