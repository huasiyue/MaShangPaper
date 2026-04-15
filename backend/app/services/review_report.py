from __future__ import annotations

from app.schemas.documents import ReviewIssue, ReviewResponse
from app.services.formatters.report import FormatReport


def build_review_response(
    report: FormatReport,
    report_text: str,
    school_id: str,
    thesis_type: str,
) -> ReviewResponse:
    issues = [
        ReviewIssue(
            level=getattr(issue.level, "value", str(issue.level)),
            location=issue.location,
            description=issue.description,
            suggestion=issue.suggestion,
            current_value=issue.current_value,
            expected_value=issue.expected_value,
        )
        for issue in report.issues
    ]

    return ReviewResponse(
        filename=report.filename,
        school_id=school_id,
        thesis_type=thesis_type,
        total_issues=report.total_issues,
        error_count=report.error_count,
        warning_count=report.warning_count,
        info_count=report.info_count,
        issues=issues,
        report_text=report_text,
    )
