from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from app.schemas.documents import ReviewResponse


class BaseFormatter(ABC):
    school_id: str

    @abstractmethod
    def review_document(self, input_path: Path, thesis_type: str) -> ReviewResponse:
        raise NotImplementedError

    @abstractmethod
    def format_document(
        self,
        input_path: Path,
        output_path: Path,
        thesis_type: str,
    ) -> tuple[Path, ReviewResponse]:
        raise NotImplementedError

