"""Shared, project-scoped personal export settings (CPython and IronPython)."""
import io
import json
import os
import tempfile
import time

from WWP_settings import _appdata_root, _project_context


def _try_file_lock(path):
    """Return a release callback for a lock owned by the current process."""
    try:
        from System.IO import FileStream, FileMode, FileAccess, FileShare
    except ImportError:
        stream = open(path, "a+b")
        try:
            if os.name == "nt":
                import msvcrt
                stream.seek(0, 2)
                if not stream.tell():
                    stream.write(b"0")
                    stream.flush()
                stream.seek(0)
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            return stream.close
        except Exception:
            stream.close()
            raise
    stream = FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, getattr(FileShare, "None"))
    return stream.Dispose


class _Encoder(json.JSONEncoder):
    def default(self, obj):
        if type(obj).__name__ in ("long", "Int64", "Int32"):
            return int(obj)
        return super(_Encoder, self).default(obj)


class _LocalConfig(object):
    def __init__(self, store, namespace, fallback):
        object.__setattr__(self, "_store", store)
        object.__setattr__(self, "_namespace", namespace)
        object.__setattr__(self, "_fallback", fallback)
        object.__setattr__(self, "_settings", store.read(namespace).get("settings", {}))

    def __getattr__(self, name):
        if name in self._settings:
            return self._settings[name]
        return getattr(self._fallback, name)

    def __setattr__(self, name, value):
        self._settings[name] = value

    def save(self):
        def update(bucket):
            block = bucket.setdefault(self._namespace, {})
            block.setdefault("settings", {}).update(self._settings)
        self._store._update(update)


class ExportSettingsStore(object):
    def __init__(self, doc, path=None):
        self.path = path or os.path.join(_appdata_root(), "ExportSchedules.user.json")
        self.project = _project_context(doc)
        self.mode = self._bucket(self._load()).get("storage_mode", "local")

    def _load(self):
        if not os.path.isfile(self.path):
            return {"schema_version": 1, "projects": {}}
        with io.open(self.path, "r", encoding="utf-8") as stream:
            data = json.load(stream)
        if (not isinstance(data, dict) or data.get("schema_version") != 1
                or not isinstance(data.get("projects"), dict)):
            raise ValueError("Unsupported export settings file: " + self.path)
        return data

    def _bucket(self, data):
        bucket = data["projects"].get(self.project["key"], {})
        if not isinstance(bucket, dict):
            raise ValueError("Invalid project settings in " + self.path)
        return bucket

    def read(self, namespace):
        data = self._bucket(self._load()).get(namespace, {})
        if not isinstance(data, dict):
            raise ValueError("Invalid export settings in " + self.path)
        return data

    def _update(self, update):
        folder = os.path.dirname(self.path)
        if not os.path.isdir(folder):
            try:
                os.makedirs(folder)
            except OSError:
                if not os.path.isdir(folder):
                    raise
        # Multiple Revit sessions can share this file. Lock the read/modify/write
        # and replace atomically so another tool's latest settings survive.
        for attempt in range(40):
            try:
                release = _try_file_lock(self.path + ".lock")
                break
            except Exception:
                if attempt == 39:
                    raise IOError("Export settings are busy or unavailable: " + self.path)
                time.sleep(0.05)
        temporary = None
        try:
            data = self._load()
            bucket = self._bucket(data)
            data["projects"][self.project["key"]] = bucket
            bucket["project_identity"] = self.project["identity"]
            bucket["project_label"] = self.project["label"]
            update(bucket)
            raw = json.dumps(data, indent=2, ensure_ascii=True, cls=_Encoder)
            handle, temporary = tempfile.mkstemp(prefix="Export2Ex-", suffix=".tmp", dir=folder)
            os.close(handle)
            with io.open(temporary, "w", encoding="utf-8") as stream:
                stream.write(u"{}".format(raw))
            if hasattr(os, "replace"):
                os.replace(temporary, self.path)
            elif os.path.exists(self.path):
                from System.IO import File
                File.Replace(temporary, self.path, None)
            else:
                os.rename(temporary, self.path)
            temporary = None
        finally:
            try:
                if temporary and os.path.isfile(temporary):
                    os.remove(temporary)
            finally:
                release()

    def write(self, namespace, value):
        self._update(lambda bucket: bucket.__setitem__(namespace, value))

    def get_config(self, namespace, fallback):
        config = _LocalConfig(self, namespace, fallback)
        return config, config.save

    def choose(self, ui):
        modes = [self.mode, "project" if self.mode == "local" else "local"]
        labels = {
            "local": "Use my local settings (JSON in my user folder)",
            "project": "Use shared project settings (Project Information parameter)",
        }
        selected = ui.uiUtils_select_indices(
            [labels[mode] for mode in modes],
            title="Export Schedules - Settings Storage",
            prompt=("Choose where to load and save settings for this project.\n"
                    "Local settings use one personal JSON file, separated by project.\n"
                    "The last choice is listed first. Existing sets stay in their current location.\n\n"
                    "Local file: " + self.path),
            multiselect=False,
        )
        if not selected:
            return False
        self.mode = modes[selected[0]]
        self._update(lambda bucket: bucket.__setitem__("storage_mode", self.mode))
        return True
