from dataclasses import dataclass


@dataclass(frozen=True)
class YZUSchoolConfig:
    school_id: str = "yzu"
    display_name: str = "扬州大学"
    default_thesis_type: str = "thesis"
    supported_thesis_types: tuple[str, ...] = ("thesis", "design_report")
    header_text: str = "扬州大学本科生毕业论文"


YZU_CONFIG = YZUSchoolConfig()

