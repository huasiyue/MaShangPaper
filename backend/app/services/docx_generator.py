from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.integrations.legacy_scripts import convert_markdown_to_word
from app.services.assets import rewrite_asset_urls_to_local_paths
from app.services.markdown_parser import DocumentModel, parse_markdown


@dataclass(slots=True)
class GenerationResult:
    output_path: Path
    markdown_path: Path
    document: DocumentModel


class DocxGenerator:
    def generate(self, content: str, markdown_path: Path, output_path: Path,
                 school_id: str = "sdfmu", thesis_type: str = "thesis") -> GenerationResult:
        processed_content = rewrite_asset_urls_to_local_paths(content)
        markdown_path.write_text(processed_content, encoding="utf-8")
        document = parse_markdown(content)
        convert_markdown_to_word(markdown_path, output_path, school_id, thesis_type)
        return GenerationResult(
            output_path=output_path,
            markdown_path=markdown_path,
            document=document,
        )
