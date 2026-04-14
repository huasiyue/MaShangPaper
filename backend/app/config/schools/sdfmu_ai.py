from dataclasses import dataclass


@dataclass(frozen=True)
class SDFMUAISchoolConfig:
    school_id: str = "sdfmu_ai"
    display_name: str = "山东第一医科大学-医学信息与人工智能学院"
    default_thesis_type: str = "thesis"
    supported_thesis_types: tuple[str, ...] = ("thesis", "design_report")
    header_text: str = "山东第一医科大学 医学信息与人工智能学院"


SDFMU_AI_CONFIG = SDFMUAISchoolConfig()
