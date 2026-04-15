#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学校格式化器注册表。
每个学校模块在导入时调用 register_formatter() 注册自身。
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from app.services.formatters.base import BaseFormatter

_FORMATTERS: dict[str, type[BaseFormatter]] = {}


def register_formatter(school_id: str, formatter_cls: type[BaseFormatter]) -> None:
    _FORMATTERS[school_id] = formatter_cls


def get_formatter_class(school_id: str) -> type[BaseFormatter]:
    if school_id not in _FORMATTERS:
        raise ValueError(f"未注册的学校模板: {school_id}")
    return _FORMATTERS[school_id]


def list_supported_schools() -> tuple[str, ...]:
    return tuple(_FORMATTERS.keys())


# 导入各学校模块以触发注册
from app.services.formatters.schools import yzu as _yzu  # noqa: F401, E402
