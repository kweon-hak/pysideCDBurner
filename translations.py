from __future__ import annotations

from pathlib import Path

# Default language display names. New languages discovered in locales/*.ini
# will be added automatically if they are not present here.
DEFAULT_LANGUAGE_NAMES = {"en": "English"}

LOCALES_DIR = Path(__file__).parent / "locales"
META_SECTION = "meta"
TRANSLATION_SECTION = "translations"

# English text is used as the lookup key; values override it when the UI
# language is switched.
DEFAULT_TRANSLATIONS: dict[str, dict[str, str]] = {}


def _unescape_value(value: str) -> str:
    """Convert simple escape sequences into their runtime form."""
    value = value.replace("\\\\", "\\")
    value = value.replace("\\n", "\n")
    value = value.replace("\\t", "\t")
    return value


def _parse_ini_file(path: Path) -> tuple[str, dict[str, str], str | None]:
    lang_code = path.stem
    lang_name: str | None = None
    translations: dict[str, str] = {}
    current_section: str | None = None
    for raw in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        stripped = raw.strip()
        if not stripped or stripped.startswith(("#", ";")):
            continue
        if stripped.startswith("[") and stripped.endswith("]"):
            current_section = stripped[1:-1].strip()
            continue
        if "=" not in raw:
            continue
        key, value = raw.split("=", 1)
        key = key.rstrip()
        value = value.lstrip()
        section = current_section or TRANSLATION_SECTION
        if section == META_SECTION:
            if key == "code" and value:
                lang_code = value.strip() or lang_code
            elif key == "name" and value:
                lang_name = value.strip()
            continue
        if section != TRANSLATION_SECTION:
            continue
        translations[key] = _unescape_value(value)
    return lang_code, translations, lang_name


def _load_translations() -> tuple[dict[str, dict[str, str]], dict[str, str]]:
    translations = {lang: table.copy() for lang, table in DEFAULT_TRANSLATIONS.items()}
    language_names = dict(DEFAULT_LANGUAGE_NAMES)
    if LOCALES_DIR.is_dir():
        for path in sorted(LOCALES_DIR.glob("*.ini")):
            try:
                lang_code, parsed, lang_name = _parse_ini_file(path)
            except Exception:
                continue
            merged = translations.get(lang_code, {}).copy()
            merged.update(parsed)
            translations[lang_code] = merged
            if lang_name:
                language_names.setdefault(lang_code, lang_name)
    for lang in translations:
        language_names.setdefault(lang, lang)
    return translations, language_names


TRANSLATIONS, LANGUAGE_NAMES = _load_translations()
