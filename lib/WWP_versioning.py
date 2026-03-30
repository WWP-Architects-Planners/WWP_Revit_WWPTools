import json
import os


_VERSION_CACHE = None


def _version_file_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "WWPTools.version.json")


def get_installed_version(default="dev"):
    global _VERSION_CACHE
    if _VERSION_CACHE is not None:
        return _VERSION_CACHE

    version = ""
    try:
        with open(_version_file_path(), "r") as version_file:
            data = json.load(version_file)
        if isinstance(data, dict):
            version = str(data.get("version") or "").strip()
    except Exception:
        version = ""

    _VERSION_CACHE = version or str(default or "dev")
    return _VERSION_CACHE


def format_window_title(title, version=None):
    base_title = str(title or "WWPTools").strip() or "WWPTools"
    version_text = str(version or get_installed_version("")).strip()
    if not version_text:
        return base_title

    suffix = " v{}".format(version_text)
    if suffix.lower() in base_title.lower():
        return base_title
    return "{}{}".format(base_title, suffix)


def apply_window_title(window, title=None):
    if window is None:
        return ""

    try:
        current_title = title if title is not None else getattr(window, "Title", "")
    except Exception:
        current_title = title or ""

    formatted_title = format_window_title(current_title)
    try:
        window.Title = formatted_title
    except Exception:
        pass
    return formatted_title
