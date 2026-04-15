#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文格式审查报告领域类型。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class FormatIssueLevel(str, Enum):
    ERROR = "错误"
    WARNING = "警告"
    INFO = "提示"


@dataclass
class FormatIssue:
    level: FormatIssueLevel
    location: str
    description: str
    suggestion: str = ""
    current_value: str = ""
    expected_value: str = ""


@dataclass
class FormatReport:
    filename: str
    total_issues: int = 0
    error_count: int = 0
    warning_count: int = 0
    info_count: int = 0
    issues: list[FormatIssue] = field(default_factory=list)
    summary: str = ""

    def add_issue(
        self,
        level: FormatIssueLevel,
        location: str,
        description: str,
        suggestion: str = "",
        current_value: str = "",
        expected_value: str = "",
    ) -> None:
        issue = FormatIssue(
            level=level,
            location=location,
            description=description,
            suggestion=suggestion,
            current_value=current_value,
            expected_value=expected_value,
        )
        self.issues.append(issue)
        self.total_issues += 1
        if level == FormatIssueLevel.ERROR:
            self.error_count += 1
        elif level == FormatIssueLevel.WARNING:
            self.warning_count += 1
        else:
            self.info_count += 1
