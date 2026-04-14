from __future__ import annotations

from pathlib import Path

from app.integrations.legacy_scripts import build_yzu_formatter
from app.schemas.documents import ReviewResponse
from app.services.formatters.base import BaseFormatter
from app.services.review_report import build_review_response


class YZUFormatter(BaseFormatter):
    school_id = "yzu"

    def review_document(self, input_path: Path, thesis_type: str) -> ReviewResponse:
        formatter = build_yzu_formatter(input_path=input_path, thesis_type=thesis_type)
        if not formatter.load_document():
            raise ValueError(f"无法加载文档: {input_path.name}")

        formatter.process_document()
        return build_review_response(formatter, school_id=self.school_id, thesis_type=thesis_type)

    def format_document(
        self,
        input_path: Path,
        output_path: Path,
        thesis_type: str,
    ) -> tuple[Path, ReviewResponse]:
        formatter = build_yzu_formatter(
            input_path=input_path,
            output_path=output_path,
            thesis_type=thesis_type,
        )
        if not formatter.load_document():
            raise ValueError(f"无法加载文档: {input_path.name}")

        formatter.process_document()
        if not formatter.save_document():
            raise ValueError(f"无法保存格式化文档: {output_path.name}")

        review = build_review_response(formatter, school_id=self.school_id, thesis_type=thesis_type)
        return Path(formatter.output_path), review

