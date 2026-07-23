from pathlib import Path


BACKEND_DIR = Path(__file__).resolve().parent.parent
LEGACY_STANDARDS_DIR = BACKEND_DIR / "standards"


def category_standard_path(category: str, filename: str, subdir: str | None = None) -> str:
    category_dir = BACKEND_DIR / str(category or "").lower() / "standards"
    preferred = category_dir / filename
    if subdir:
        preferred = category_dir / subdir / filename
    if preferred.exists():
        return str(preferred)

    legacy = LEGACY_STANDARDS_DIR / filename
    if subdir:
        legacy = LEGACY_STANDARDS_DIR / subdir / filename
    if legacy.exists():
        return str(legacy)

    return str(preferred)


def category_standards_dir(category: str, fallback_to_legacy: bool = False) -> str:
    category_dir = BACKEND_DIR / str(category or "").lower() / "standards"
    if category_dir.exists() or not fallback_to_legacy:
        return str(category_dir)
    return str(LEGACY_STANDARDS_DIR)
