from __future__ import annotations

import importlib.util
import sys
from functools import lru_cache
from pathlib import Path

from app.core.settings import BACKEND_DIR


def _load_module(module_name: str, file_path: Path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load module from {file_path}")

    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


@lru_cache(maxsize=1)
def get_convert_module():
    """加载带公式支持的转换器（backend 版本）。"""
    return _load_module("word_converter", BACKEND_DIR / "app" / "services" / "word_converter.py")


def convert_markdown_to_word(markdown_path: Path, output_path: Path,
                            school_id: str = "yzu", thesis_type: str = "thesis") -> None:
    module = get_convert_module()
    module.convert_markdown_to_word(str(markdown_path), str(output_path), school_id, thesis_type)
