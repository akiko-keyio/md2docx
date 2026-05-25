"""Style profiles for md2docx."""

from .academic_manuscript import PROFILE as ACADEMIC_MANUSCRIPT

PROFILES = {
    "academic-manuscript": ACADEMIC_MANUSCRIPT,
}


def get_profile(name: str) -> dict:
    if name not in PROFILES:
        raise ValueError(f"Unknown profile: {name}. Available: {list(PROFILES.keys())}")
    return PROFILES[name]
