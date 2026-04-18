#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学校格式化器注册表。
每个学校模块在导入时调用 register_formatter() 注册自身。
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

if TYPE_CHECKING:
    from app.services.formatters.base import BaseFormatter
    from app.services.formatters.spec import ThesisFormatSpec

_FORMATTERS: dict[str, type[BaseFormatter]] = {}
_SPEC_FACTORIES: dict[str, Callable[[str], ThesisFormatSpec]] = {}


def register_formatter(school_id: str, formatter_cls: type[BaseFormatter]) -> None:
    _FORMATTERS[school_id] = formatter_cls


def register_spec_factory(school_id: str, factory: Callable[[str], ThesisFormatSpec]) -> None:
    """注册规格工厂函数。每个学校模块调用一次，如 register_spec_factory("sdfmu", create_sdfmu_spec)。"""
    _SPEC_FACTORIES[school_id] = factory


def get_formatter_class(school_id: str) -> type[BaseFormatter]:
    if school_id not in _FORMATTERS:
        raise ValueError(f"未注册的学校模板: {school_id}")
    return _FORMATTERS[school_id]


def get_spec(school_id: str, thesis_type: str) -> "ThesisFormatSpec":
    """根据 school_id 自动查找已注册的规格工厂并生成 spec。"""
    if school_id not in _SPEC_FACTORIES:
        raise ValueError(f"未注册的学校规格: {school_id}")
    return _SPEC_FACTORIES[school_id](thesis_type)


def list_supported_schools() -> tuple[str, ...]:
    return tuple(_FORMATTERS.keys())


# 导入各学校模块以触发注册
from app.services.formatters.schools import sdfmu as _sdfmu  # noqa: F401, E402
