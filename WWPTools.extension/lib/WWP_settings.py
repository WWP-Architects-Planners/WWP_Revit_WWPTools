import hashlib
import json
import os
import re
import time

try:
    from pyrevit import DB
except Exception:
    DB = None


SCHEMA_VERSION = 1


def _appdata_root():
    root = os.environ.get("APPDATA")
    if not root:
        root = os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    return os.path.join(root, "pyRevit", "WWPTools")


def _ensure_dir(path):
    if path and not os.path.isdir(path):
        os.makedirs(path)


def _sanitize_file_name(value):
    text = re.sub(r'[<>:"/\\|?*]+', "_", str(value or "").strip())
    return text or "ToolSettings"


def _tool_settings_path(tool_name):
    return os.path.join(_appdata_root(), "{}.settings.json".format(_sanitize_file_name(tool_name)))


def _read_json_file(path):
    if not path or not os.path.isfile(path):
        return {}
    try:
        with open(path, "r") as fp:
            data = json.load(fp)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _json_safe(value):
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    if isinstance(value, dict):
        return {str(key): _json_safe(val) for key, val in value.items()}
    return str(value)


def _safe_getattr(obj, name, default=None):
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _read_legacy_value(source, name):
    if source is None:
        return False, None
    if isinstance(source, dict):
        if name in source:
            return True, source[name]
        return False, None
    try:
        return True, getattr(source, name)
    except AttributeError:
        return False, None
    except Exception:
        return False, None


def _safe_model_path_to_string(model_path):
    if model_path is None:
        return ""
    if DB is not None:
        try:
            return DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
        except Exception:
            pass
    for member_name in ("ToUserVisiblePath", "ToString"):
        member = getattr(model_path, member_name, None)
        if callable(member):
            try:
                value = member()
                if value:
                    return str(value)
            except Exception:
                continue
    return str(model_path)


def _project_context(doc):
    title = _safe_getattr(doc, "Title", "") or "Active Document"
    label = title
    candidates = []

    fingerprint = getattr(doc, "GetFingerprintGUID", None)
    if callable(fingerprint):
        try:
            value = fingerprint()
            if value:
                candidates.append("fingerprint:" + str(value))
        except Exception:
            pass

    try:
        value = getattr(doc, "WorksharingCentralGUID", None)
        if value:
            candidates.append("central-guid:" + str(value))
    except Exception:
        pass

    try:
        if _safe_getattr(doc, "IsWorkshared", False):
            central_model_path = doc.GetWorksharingCentralModelPath()
            central_path = _safe_model_path_to_string(central_model_path)
            if central_path:
                candidates.append("central-path:" + central_path)
                label = os.path.basename(central_path.rstrip("\\/")) or label
    except Exception:
        pass

    cloud_model_path_getter = getattr(doc, "GetCloudModelPath", None)
    if callable(cloud_model_path_getter):
        try:
            cloud_path = _safe_model_path_to_string(cloud_model_path_getter())
            if cloud_path:
                candidates.append("cloud-path:" + cloud_path)
                label = os.path.basename(cloud_path.rstrip("\\/")) or label
        except Exception:
            pass

    try:
        path_name = getattr(doc, "PathName", "") or ""
        if path_name:
            candidates.append("path:" + path_name)
            label = os.path.basename(path_name.rstrip("\\/")) or label
    except Exception:
        pass

    if title:
        candidates.append("title:" + title)

    identity = ""
    for candidate in candidates:
        if candidate:
            identity = candidate
            break
    if not identity:
        identity = "title:" + title

    key = hashlib.sha1(identity.lower().encode("utf-8")).hexdigest()[:16]
    return {
        "key": key,
        "identity": identity,
        "label": label or title or "Active Document",
    }


class ProjectToolSettings(object):
    def __init__(self, tool_name, doc=None, legacy_sources=None, legacy_file_paths=None):
        object.__setattr__(self, "_tool_name", tool_name)
        object.__setattr__(self, "_file_path", _tool_settings_path(tool_name))
        object.__setattr__(self, "_project_context", _project_context(doc))
        object.__setattr__(self, "_legacy_sources", list(legacy_sources or []))

        payload = _read_json_file(self._file_path)
        legacy_payload = {}
        if not isinstance(payload.get("projects"), dict):
            legacy_payload.update(payload)
            payload = {}

        for legacy_file_path in legacy_file_paths or []:
            legacy_data = _read_json_file(legacy_file_path)
            if isinstance(legacy_data.get("projects"), dict):
                continue
            legacy_payload.update(legacy_data)

        data = {
            "schema_version": SCHEMA_VERSION,
            "tool_name": tool_name,
            "last_project_key": "",
            "projects": {},
        }
        data.update(payload)
        if not isinstance(data.get("projects"), dict):
            data["projects"] = {}

        project_key = self._project_context["key"]
        project_bucket = data["projects"].get(project_key)
        if not isinstance(project_bucket, dict):
            project_bucket = {}
            data["projects"][project_key] = project_bucket
        if not isinstance(project_bucket.get("settings"), dict):
            project_bucket["settings"] = {}

        if not project_bucket["settings"] and legacy_payload:
            project_bucket["settings"].update(_json_safe(legacy_payload))

        project_bucket["project_identity"] = self._project_context["identity"]
        project_bucket["project_label"] = self._project_context["label"]

        object.__setattr__(self, "_data", data)
        object.__setattr__(self, "_project_bucket", project_bucket)
        object.__setattr__(self, "_settings", project_bucket["settings"])

    @property
    def file_path(self):
        return self._file_path

    @property
    def project_key(self):
        return self._project_context["key"]

    @property
    def project_label(self):
        return self._project_context["label"]

    def __getattr__(self, name):
        if name in self._settings:
            return self._settings[name]
        for source in self._legacy_sources:
            found, value = _read_legacy_value(source, name)
            if found:
                return value
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        self._settings[name] = _json_safe(value)

    def save(self):
        _ensure_dir(os.path.dirname(self._file_path))
        self._project_bucket["project_identity"] = self._project_context["identity"]
        self._project_bucket["project_label"] = self._project_context["label"]
        self._project_bucket["last_used_utc"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
        self._data["last_project_key"] = self._project_context["key"]
        self._data["schema_version"] = SCHEMA_VERSION
        self._data["tool_name"] = self._tool_name
        with open(self._file_path, "w") as fp:
            json.dump(self._data, fp, indent=2, sort_keys=True)


def get_tool_settings(tool_name, doc=None, legacy_sources=None, legacy_file_paths=None):
    settings = ProjectToolSettings(
        tool_name=tool_name,
        doc=doc,
        legacy_sources=legacy_sources,
        legacy_file_paths=legacy_file_paths,
    )
    return settings, settings.save
