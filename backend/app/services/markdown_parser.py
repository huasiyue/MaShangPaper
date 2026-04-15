from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
import re


class BlockType(str, Enum):
    HEADING = "heading"
    PARAGRAPH = "paragraph"
    TABLE = "table"
    IMAGE = "image"
    LIST = "list"
    QUOTE = "quote"
    CODE = "code"
    REFERENCE = "reference"
    EQUATION = "equation"


@dataclass(slots=True)
class DocumentBlock:
    block_type: BlockType
    text: str
    level: int | None = None
    lines: list[str] = field(default_factory=list)


@dataclass(slots=True)
class DocumentModel:
    raw_text: str
    title: str | None
    blocks: list[DocumentBlock]

    @property
    def heading_count(self) -> int:
        return sum(1 for block in self.blocks if block.block_type == BlockType.HEADING)

    @property
    def table_count(self) -> int:
        return sum(1 for block in self.blocks if block.block_type == BlockType.TABLE)

    @property
    def image_count(self) -> int:
        return sum(1 for block in self.blocks if block.block_type == BlockType.IMAGE)

    @property
    def reference_count(self) -> int:
        return sum(1 for block in self.blocks if block.block_type == BlockType.REFERENCE)


def parse_markdown(content: str) -> DocumentModel:
    lines = content.splitlines()
    blocks: list[DocumentBlock] = []
    title: str | None = None

    in_code_block = False
    in_math_block = False
    math_buffer: list[str] = []
    table_buffer: list[str] = []

    def flush_table() -> None:
        nonlocal table_buffer
        if table_buffer:
            blocks.append(
                DocumentBlock(
                    block_type=BlockType.TABLE,
                    text="\n".join(table_buffer),
                    lines=table_buffer[:],
                )
            )
            table_buffer = []

    def flush_math() -> None:
        nonlocal math_buffer
        if math_buffer:
            blocks.append(
                DocumentBlock(
                    block_type=BlockType.EQUATION,
                    text="\n".join(math_buffer),
                    lines=math_buffer[:],
                )
            )
            math_buffer = []

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()

        # 代码块
        if stripped.startswith("```"):
            flush_table()
            in_code_block = not in_code_block
            blocks.append(DocumentBlock(block_type=BlockType.CODE, text=stripped))
            continue

        if in_code_block:
            blocks.append(DocumentBlock(block_type=BlockType.CODE, text=line))
            continue

        # 数学公式块 $$...$$
        if stripped.startswith("$$"):
            flush_table()
            # 单行公式：$$ E = mc^2 $$
            if stripped.endswith("$$") and len(stripped) > 4 and not in_math_block:
                blocks.append(
                    DocumentBlock(
                        block_type=BlockType.EQUATION,
                        text=stripped[2:-2].strip(),
                    )
                )
                continue
            # 多行公式块开始/结束
            if not in_math_block:
                in_math_block = True
                math_buffer = []
            else:
                in_math_block = False
                flush_math()
            continue

        if in_math_block:
            math_buffer.append(stripped)
            continue

        if not stripped:
            flush_table()
            continue

        if stripped.startswith("|"):
            table_buffer.append(stripped)
            continue

        flush_table()

        if stripped.startswith("#"):
            level = len(stripped) - len(stripped.lstrip("#"))
            text = stripped[level:].strip()
            if title is None and level == 1:
                title = text
            blocks.append(DocumentBlock(block_type=BlockType.HEADING, text=text, level=level))
        elif stripped.startswith("![" ):
            blocks.append(DocumentBlock(block_type=BlockType.IMAGE, text=stripped))
        elif stripped.startswith(">"):
            blocks.append(DocumentBlock(block_type=BlockType.QUOTE, text=stripped[1:].strip()))
        elif stripped.startswith(("-", "*", "+")) or re.match(r"^\d+\.", stripped):
            blocks.append(DocumentBlock(block_type=BlockType.LIST, text=stripped))
        elif stripped.startswith("[") and "]" in stripped:
            blocks.append(DocumentBlock(block_type=BlockType.REFERENCE, text=stripped))
        else:
            blocks.append(DocumentBlock(block_type=BlockType.PARAGRAPH, text=stripped))

    flush_table()

    if title is None:
        for block in blocks:
            if block.block_type in {BlockType.HEADING, BlockType.PARAGRAPH} and block.text:
                title = block.text
                break

    return DocumentModel(raw_text=content, title=title, blocks=blocks)
