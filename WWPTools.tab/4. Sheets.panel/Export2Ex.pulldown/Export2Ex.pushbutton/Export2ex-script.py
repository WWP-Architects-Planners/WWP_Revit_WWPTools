import csv
import importlib
import io
import json
import os
import re
import shutil
import sys
import tempfile
import time
import traceback

import clr
clr.AddReference('System.Xml')
from System import String
from System.Collections.Generic import List
from System.IO import File, StreamReader, StreamWriter
from System.Text import Encoding, UTF8Encoding

from pyrevit import DB
from WWP_settings import get_tool_settings
from WWP_versioning import apply_window_title



CONFIG_LAST_EXCEL_PATH = "last_excel_path"
CONFIG_LAST_CSV_DIR = "last_csv_dir"
CONFIG_LAST_SCHEDULE_IDS = "last_schedule_ids"
CONFIG_LAST_CSV_MODE = "last_csv_mode"
CONFIG_LAST_CSV_DELIM = "last_csv_delim"
CONFIG_LAST_EXPORT_MODE = "last_export_mode"
CONFIG_LAST_CSV_EXPORT_TITLE = "last_csv_export_title"
CONFIG_LAST_CSV_COLUMN_HEADERS = "last_csv_column_headers"
CONFIG_LAST_CSV_GROUP_HEADERS = "last_csv_group_headers"
CONFIG_LAST_CSV_GROUPED_COLUMN_HEADERS = "last_csv_grouped_column_headers"
CONFIG_LAST_CSV_TEXT_QUALIFIER = "last_csv_text_qualifier"
CONFIG_LAST_USE_CATEGORY_SHEET_NAME = "last_use_category_sheet_name"
LOG_FILE_NAME = "Export2Ex.log"
ALLOWED_EXCEL_EXTENSIONS = (".xlsx", ".xlsm")
PARAM_SAVED_SETS = "! P_STATS_Export_Text"
SAVED_SET_NAMESPACE = "export2ex"




def _elem_id_int(eid):
    try:
        return int(eid.Value)      # Revit 2024+
    except AttributeError:
        return int(eid.Value)  # Revit 2023-

def sanitize_sheet_name(name):
    invalid = r"[:\\/?*\[\]]"
    safe = re.sub(invalid, "_", name)
    safe = safe.strip()
    if not safe:
        safe = "Schedule"
    return safe[:31]


def sanitize_file_name(name):
    invalid = r'[<>:"/\\|?*]'
    safe = re.sub(invalid, "_", name).strip()
    return safe or "Schedule"


def _pluralize(name):
    """Return a simple English plural of name (for sheet naming). Most Revit category
    names are already plural (Walls, Areas, Rooms); this handles the singular ones."""
    if not name:
        return name
    lower = name.lower()
    # Already ends in s/x/z -- treat as plural (Walls, Areas, Rooms, Stairs, ...)
    if lower[-1] in "sxz":
        return name
    # Ends in consonant + y: change to ies (e.g. "Category" -> "Categories")
    if lower.endswith("y") and len(name) > 1 and lower[-2] not in "aeiou":
        return name[:-1] + "ies"
    # Ends in ch/sh: add es
    if lower.endswith("ch") or lower.endswith("sh"):
        return name + "es"
    # Default: add s  (Parking -> Parkings, Fixture -> Fixtures, etc.)
    return name + "s"


def _get_schedule_category_name(doc, view):
    """Return the Revit category name for a schedule view, falling back to view.Name."""
    try:
        cat_id = view.Definition.CategoryId
        if cat_id is None:
            return view.Name
        try:
            cat = DB.Category.GetCategory(doc, cat_id)
            if cat is not None:
                return cat.Name
        except Exception:
            pass
        for cat in doc.Settings.Categories:
            try:
                if cat.Id.IntegerValue == cat_id.IntegerValue:
                    return cat.Name
            except Exception:
                continue
    except Exception:
        pass
    return view.Name


def normalize_excel_output_path(path, default_ext=".xlsx"):
    value = os.path.expandvars((path or "").strip())
    if not value:
        return ""
    root, ext = os.path.splitext(value)
    if not ext:
        return value + default_ext
    if ext.lower() in ALLOWED_EXCEL_EXTENSIONS:
        return value
    return ""


def _pick_save_file(title, filter_text, default_extension, initial_directory, file_name):
    clr.AddReference("PresentationFramework")
    from Microsoft.Win32 import SaveFileDialog

    dialog = SaveFileDialog()
    dialog.Title = title or "Save File"
    dialog.Filter = filter_text or "All files (*.*)|*.*"
    if default_extension:
        dialog.DefaultExt = default_extension
        dialog.AddExtension = True
    if initial_directory:
        initial_directory = os.path.expandvars(initial_directory)
    if initial_directory and os.path.isdir(initial_directory):
        dialog.InitialDirectory = initial_directory
    if file_name:
        dialog.FileName = file_name
    result = dialog.ShowDialog()
    if result:
        return dialog.FileName
    return None


def _pick_folder(title, initial_directory):
    clr.AddReference("System.Windows.Forms")
    from System.Windows.Forms import FolderBrowserDialog, DialogResult

    dialog = FolderBrowserDialog()
    dialog.Description = title or "Select Folder"
    if initial_directory:
        initial_directory = os.path.expandvars(initial_directory)
    if initial_directory and os.path.isdir(initial_directory):
        dialog.SelectedPath = initial_directory
    result = dialog.ShowDialog()
    if result == DialogResult.OK:
        return dialog.SelectedPath
    return None


def get_active_doc():
    """Resolve current document without importing pyrevit.revit (CPython 6.1 safety)."""
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc:
            return uidoc.Document
    except Exception:
        pass
    return None


def read_saved_sets(doc, param_name=None):
    try:
        proj_info = doc.ProjectInformation
        if proj_info is None:
            return {}
        param = proj_info.LookupParameter(param_name or PARAM_SAVED_SETS)
        if param is None:
            return {}
        raw = (param.AsString() or "").strip()
        if not raw:
            return {}
        data = json.loads(raw)
        if not isinstance(data, dict):
            return {}
        if isinstance(data.get(SAVED_SET_NAMESPACE), dict):
            return data.get(SAVED_SET_NAMESPACE) or {}
        if any(k in data for k in _ALL_NAMESPACES):
            return {}
        return data
    except Exception:
        return {}


_ALL_NAMESPACES = ("export2ex", "export2ex_beta", "mass_stats")


def _looks_like_legacy_saved_sets(data):
    if not isinstance(data, dict) or not data:
        return False
    if any(k in data for k in _ALL_NAMESPACES):
        return False
    return all(isinstance(v, dict) for v in data.values())


def write_saved_sets(doc, sets_dict, param_name=None):
    try:
        proj_info = doc.ProjectInformation
        if proj_info is None:
            return False
        param = proj_info.LookupParameter(param_name or PARAM_SAVED_SETS)
        t = DB.Transaction(doc, "Save Export2Ex Settings")
        t.Start()
        try:
            if param is None:
                t.RollBack()
                return False
            raw = (param.AsString() or "").strip()
            try:
                payload = json.loads(raw) if raw else {}
            except Exception:
                payload = {}
            if not isinstance(payload, dict):
                payload = {}
            if _looks_like_legacy_saved_sets(payload):
                payload = {SAVED_SET_NAMESPACE: payload}
            payload[SAVED_SET_NAMESPACE] = sets_dict
            param.Set(json.dumps(payload, ensure_ascii=False, indent=2))
            t.Commit()
            return True
        except Exception as inner:
            try:
                t.RollBack()
            except Exception:
                pass
            log_exception("write_saved_sets transaction", inner)
            return False
    except Exception as exc:
        log_exception("write_saved_sets", exc)
        return False


def _get_proj_info_text_params(doc):
    """Return sorted names of writable text parameters on Project Information."""
    try:
        proj_info = doc.ProjectInformation
        if proj_info is None:
            return []
        names = []
        for param in proj_info.Parameters:
            try:
                if param.StorageType == DB.StorageType.String:
                    names.append(param.Definition.Name)
            except Exception:
                continue
        return sorted(set(names))
    except Exception:
        return []


def _ensure_saved_sets_param(doc, config, ui, save_config):
    """Return the Project Information param name to use for saved sets.

    Priority: canonical param -> config-stored override -> user prompt.
    Saves the user's choice to per-project config. Returns None if unavailable.
    """
    try:
        proj_info = doc.ProjectInformation
    except Exception:
        return None
    if proj_info is None:
        return None

    if proj_info.LookupParameter(PARAM_SAVED_SETS) is not None:
        return PARAM_SAVED_SETS

    stored = config_get(config, "saved_sets_param_name", "") or ""
    if stored and proj_info.LookupParameter(stored) is not None:
        return stored

    text_params = _get_proj_info_text_params(doc)
    if not text_params:
        ui.uiUtils_alert(
            "Parameter '{}' was not found in Project Information and no other "
            "writable text parameters exist.\n\n"
            "Add a text parameter to Project Information to enable saved sets.".format(PARAM_SAVED_SETS),
            title="Export Schedules - Saved Sets",
        )
        return None

    indices = ui.uiUtils_select_indices(
        text_params,
        title="Select Parameter for Saved Sets",
        prompt=(
            "'{}' was not found in this project's Project Information.\n\n"
            "Select an existing text parameter to store saved sets:"
        ).format(PARAM_SAVED_SETS),
        multiselect=False,
    )
    if not indices:
        return None

    chosen = text_params[indices[0]]
    config.saved_sets_param_name = chosen
    if save_config:
        save_config()
    return chosen


def _normalize_namespace_data(data):
    if not isinstance(data, dict):
        return {}
    if "settings" in data or "sets" in data:
        return data
    if _looks_like_legacy_saved_sets(data):
        return {"sets": data}
    return {}


class _LocalConfig(object):
    def __init__(self, file_path, data):
        object.__setattr__(self, "_file_path", file_path)
        object.__setattr__(self, "_data", data or {})

    def __getattr__(self, name):
        if name in self._data:
            return self._data[name]
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name.startswith("_"):
            object.__setattr__(self, name, value)
        else:
            self._data[name] = value

    def save(self):
        folder = os.path.dirname(self._file_path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        with open(self._file_path, "w") as fp:
            json.dump(self._data, fp, indent=2)


class _CombinedConfig(object):
    def __init__(self, primary, primary_saver, fallback, fallback_saver):
        object.__setattr__(self, "_primary", primary)
        object.__setattr__(self, "_primary_saver", primary_saver)
        object.__setattr__(self, "_fallback", fallback)
        object.__setattr__(self, "_fallback_saver", fallback_saver)

    def __getattr__(self, name):
        for cfg in (self._primary, self._fallback):
            if cfg is None:
                continue
            try:
                return getattr(cfg, name)
            except AttributeError:
                continue
            except Exception:
                continue
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        for cfg in (self._primary, self._fallback):
            if cfg is None:
                continue
            try:
                setattr(cfg, name, value)
            except Exception:
                continue

    def save(self):
        saved = False
        errors = []
        for saver in (self._primary_saver, self._fallback_saver):
            if not callable(saver):
                continue
            try:
                saver()
                saved = True
            except Exception as exc:
                errors.append(exc)
        if not saved and errors:
            raise errors[0]


def _legacy_local_config_path():
    appdata = os.environ.get("APPDATA")
    if not appdata:
        appdata = os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    return os.path.join(appdata, "pyRevit", "WWPTools", "Export2Ex.config.json")


def _log_file_path():
    appdata = os.environ.get("APPDATA")
    if not appdata:
        appdata = os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    return os.path.join(appdata, "pyRevit", "WWPTools", "Logs", LOG_FILE_NAME)


def log_message(message):
    try:
        log_path = _log_file_path()
        folder = os.path.dirname(log_path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a") as fp:
            fp.write("[{}] {}\n".format(timestamp, message))
    except Exception:
        pass


def log_exception(context, exc):
    try:
        detail = traceback.format_exc()
    except Exception:
        detail = str(exc)
    log_message("{}: {}\n{}".format(context, str(exc), detail))


def describe_file_state(path):
    try:
        exists = os.path.exists(path)
    except Exception:
        exists = False
    try:
        is_file = os.path.isfile(path)
    except Exception:
        is_file = False
    try:
        size = os.path.getsize(path) if is_file else -1
    except Exception:
        size = -1
    return "path='{}' exists={} is_file={} size={}".format(path, exists, is_file, size)


def read_log_tail(max_chars=5000):
    log_path = _log_file_path()
    if not os.path.isfile(log_path):
        return "Log file not found:\n{}".format(log_path)
    try:
        with open(log_path, "r") as fp:
            text = fp.read()
    except Exception as exc:
        return "Failed to read log file '{}': {}".format(log_path, str(exc))
    if len(text) <= max_chars:
        return text
    return text[-max_chars:]


def show_error_report(ex):
    log_path = _log_file_path()
    report = (
        "Export2Ex failed.\n\n"
        "Error\n{}\n\n"
        "Log File\n{}\n\n"
        "Recent Log\n{}"
    ).format(str(ex), log_path, read_log_tail())
    try:
        ui = load_uiutils()
        if hasattr(ui, "uiUtils_show_text_report"):
            ui.uiUtils_show_text_report(
                "Export2Ex Error Report",
                report,
                ok_text="Close",
                cancel_text=None,
                width=900,
                height=620,
            )
            return
        ui.uiUtils_alert(report, title="Export2Ex Error Report")
    except Exception:
        pass


def _load_local_config():
    cfg_path = _legacy_local_config_path()
    data = {}
    if os.path.isfile(cfg_path):
        try:
            with open(cfg_path, "r") as fp:
                loaded = json.load(fp)
            if isinstance(loaded, dict):
                data = loaded
        except Exception:
            data = {}
    return _LocalConfig(cfg_path, data)


def get_config_and_saver():
    """Persist to a project-aware JSON file and seed from legacy config when available."""
    legacy_sources = []
    try:
        from pyrevit import script as pyrevit_script

        legacy_sources.append(pyrevit_script.get_config())
    except Exception:
        pass
    legacy_files = [_legacy_local_config_path()]
    return get_tool_settings(
        "Export2Ex",
        doc=get_active_doc(),
        legacy_sources=legacy_sources,
        legacy_file_paths=legacy_files,
    )


def config_get(config, name, default=None):
    try:
        value = getattr(config, name)
    except Exception:
        return default
    return default if value is None else value


def _normalize_path(path):
    """Store paths with %USERPROFILE% prefix so configs are portable across user accounts."""
    if not path:
        return path
    userprofile = os.environ.get("USERPROFILE", "")
    if userprofile and path.lower().startswith(userprofile.lower()):
        return "%USERPROFILE%" + path[len(userprofile):]
    return path


def _is_cloud_path(path):
    if not path:
        return False
    lower = path.lower()
    return (lower.startswith("bim 360://")
            or lower.startswith("autodesk docs://")
            or lower.startswith("autodesk forma://"))


def _docs_folder():
    try:
        from System import Environment
        return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    except Exception:
        return os.path.expanduser("~")


def get_default_dir(doc):
    if doc.IsWorkshared:
        try:
            central = doc.GetWorksharingCentralModelPath()
            if central:
                path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(central)
                if not _is_cloud_path(path):
                    return os.path.dirname(path)
        except Exception:
            pass
    if doc.PathName and not _is_cloud_path(doc.PathName):
        return os.path.dirname(doc.PathName)
    return _docs_folder()


def ensure_existing_dir(path, fallback=""):
    if path and os.path.isdir(path):
        return path
    if fallback and os.path.isdir(fallback):
        return fallback
    return ""


def collect_schedules(doc):
    schedules = []
    for view in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule):
        if view.IsTemplate:
            continue
        try:
            if view.ViewType == DB.ViewType.Legend:
                continue
        except Exception:
            pass
        if view.IsTitleblockRevisionSchedule:
            continue
        schedules.append(view)
    schedules.sort(key=lambda v: v.Name)
    return schedules


def element_id_value(elem_id):
    if elem_id is None:
        return -1
    if hasattr(elem_id, "IntegerValue"):
        return _elem_id_int(elem_id)
    if hasattr(elem_id, "Value"):
        return elem_id.Value
    try:
        return int(elem_id)
    except Exception:
        return -1


class ScheduleItem(object):
    def __init__(self, view):
        self.view = view
        display_name = "{} [id:{}]".format(view.Name, element_id_value(view.Id))
        self.display_name = display_name.replace("_", "__")


def add_lib_path():
    lib_path = os.path.join(os.path.dirname(__file__), "lib")
    if lib_path not in sys.path:
        sys.path.append(lib_path)


def load_uiutils():
    script_dir = os.path.dirname(__file__)
    lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
    if lib_path not in sys.path:
        sys.path.append(lib_path)
    import WWP_uiUtils as ui
    if not hasattr(ui, "uiUtils_select_items_with_mode"):
        try:
            ui = importlib.reload(ui)
        except Exception:
            pass
    return ui


def _export_sets_to_file(saved_sets, ui, tool_name="Export2Ex"):
    """Save current saved sets to an external .settings file."""
    clr.AddReference("PresentationFramework")
    from Microsoft.Win32 import SaveFileDialog
    dlg = SaveFileDialog()
    dlg.Title = "Save Settings File"
    dlg.Filter = "Export Settings (*.settings)|*.settings|All Files (*.*)|*.*"
    dlg.DefaultExt = "settings"
    dlg.FileName = "{}_sets.settings".format(tool_name)
    if dlg.ShowDialog() != True:
        return
    try:
        payload = {"tool": tool_name, "version": "1.0", "sets": saved_sets}
        with open(dlg.FileName, "w") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        if ui:
            ui.uiUtils_alert("Settings saved to:\n{}".format(dlg.FileName), title="Save Settings")
    except Exception as exc:
        if ui:
            ui.uiUtils_alert("Could not save settings file:\n{}".format(exc), title="Save Settings")


def _import_sets_from_file(ui):
    """Load saved sets from an external .settings file. Returns dict or None if cancelled/failed."""
    clr.AddReference("PresentationFramework")
    from Microsoft.Win32 import OpenFileDialog
    dlg = OpenFileDialog()
    dlg.Title = "Load Settings File"
    dlg.Filter = "Export Settings (*.settings)|*.settings|All Files (*.*)|*.*"
    dlg.CheckFileExists = True
    if dlg.ShowDialog() != True:
        return None
    try:
        with open(dlg.FileName, "r") as f:
            data = json.load(f)
        sets = data.get("sets", {})
        if not isinstance(sets, dict):
            raise ValueError("Invalid settings file: 'sets' key missing or not a dict.")
        return sets
    except Exception as exc:
        if ui:
            ui.uiUtils_alert("Could not load settings file:\n{}".format(exc), title="Load Settings")
        return None


def _to_net_str_list(values):
    lst = List[String]()
    for v in values:
        lst.Add(str(v))
    return lst


def _load_dialog_xaml(file_name):
    """Parse a dialog XAML that sits next to this script, after loading the shared theme."""
    clr.AddReference("PresentationFramework")
    clr.AddReference("PresentationCore")
    clr.AddReference("WindowsBase")
    from System.IO import StringReader
    from System.Windows.Markup import XamlReader
    from System.Xml import XmlReader

    xaml_path = os.path.join(os.path.dirname(__file__), file_name)
    if not os.path.isfile(xaml_path):
        raise Exception("Missing dialog XAML: {}".format(xaml_path))
    load_uiutils().uiUtils_ensure_theme()
    reader = XmlReader.Create(StringReader(File.ReadAllText(xaml_path)))
    return XamlReader.Load(reader)


def _show_export_settings_popup(current_settings, ui, owner=None):
    """Settings popup for export options. Returns updated dict or None if cancelled."""
    clr.AddReference("PresentationFramework")
    from System.Windows import WindowStartupLocation

    s = dict(current_settings)
    _ok = [False]

    win = _load_dialog_xaml("ExportSchedulesDialog.xaml")
    apply_window_title(win, "Export Settings")
    if owner is not None:
        win.Owner = owner
    else:
        win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    excel_radio = win.FindName("ExcelMode")
    csv_radio = win.FindName("CsvMode")
    chk_title = win.FindName("ExportTitle")
    chk_col_hdr = win.FindName("ExportColumnHeaders")
    chk_grpd_col = win.FindName("ExportGroupedColumnHeaders")
    chk_grp_hdr = win.FindName("ExportGroupHeaders")
    chk_use_cat = win.FindName("UseCategorySheetName")
    delim_combo = win.FindName("CsvDelimiter")
    chk_quote = win.FindName("QuoteAll")
    apply_btn = win.FindName("ApplyButton")

    if s.get("export_mode", 0) == 1:
        csv_radio.IsChecked = True
    else:
        excel_radio.IsChecked = True

    chk_title.IsChecked = bool(s.get("export_title", False))
    chk_col_hdr.IsChecked = bool(s.get("export_column_headers", True))
    chk_grpd_col.IsChecked = bool(s.get("export_grouped_column_headers", False))
    chk_grp_hdr.IsChecked = bool(s.get("export_group_headers", False))
    chk_use_cat.IsChecked = bool(s.get("use_category_sheet_name", True))

    delimiter_map = {"Comma (,)": ",", "Semicolon (;)": ";", "Tab (\\t)": "\t"}
    delim_combo.ItemsSource = _to_net_str_list(list(delimiter_map.keys()))
    cur_delim = s.get("csv_delim", ",")
    for lbl, val in delimiter_map.items():
        if val == cur_delim:
            delim_combo.SelectedItem = lbl
            break
    chk_quote.IsChecked = bool(s.get("csv_mode_quote", bool(s.get("csv_mode", 0) == 1)))

    def _apply(sv, e):
        _ok[0] = True
        win.Close()

    # CancelButton is IsCancel in the XAML: it closes the window (and so does Esc).
    apply_btn.Click += _apply
    win.ShowDialog()

    if not _ok[0]:
        return None

    s["export_mode"]                  = 1 if csv_radio.IsChecked else 0
    s["export_title"]                 = bool(chk_title.IsChecked)
    s["export_column_headers"]        = bool(chk_col_hdr.IsChecked)
    s["export_grouped_column_headers"]= bool(chk_grpd_col.IsChecked)
    s["export_group_headers"]         = bool(chk_grp_hdr.IsChecked)
    s["use_category_sheet_name"]      = bool(chk_use_cat.IsChecked)
    s["csv_mode"]                     = 1 if chk_quote.IsChecked else 0
    s["csv_delim"]                    = delimiter_map.get(str(delim_combo.SelectedItem or ""), ",")
    return s


_PROFILE_TOKEN = "%userprofile%"


def _portable_path(path):
    """Replace the current user's profile folder (C:/Users/<name>) with %userprofile% so saved
    paths, notably ACC Desktop Connector (DC) folders, resolve for every user."""
    try:
        p = (path or "").strip()
        prof = os.environ.get("USERPROFILE", "").rstrip("/" + os.sep)
        if prof and (p.lower() == prof.lower() or p.lower().startswith(prof.lower() + os.sep)):
            return _PROFILE_TOKEN + p[len(prof):]
        return p
    except Exception:
        return path


def _expand_path(path):
    """Resolve %userprofile% (and other env vars) to a real path before use."""
    try:
        return os.path.expandvars(path or "")
    except Exception:
        return path


_display_path = _portable_path


def _show_batch_dialog(saved_sets, doc, ui=None, saved_sets_param=None,
                       default_settings=None, edit_callback=None, add_callback=None):
    """Primary export UI (BatchExportDialog.xaml): one row per saved set with its schedule
    (searchable), sheet name, output path, and per-row Browse / Export / Delete.
    There is no Save button: row edits are written to Project Information once, when the
    window closes, and only if something changed."""
    clr.AddReference("PresentationFramework")
    clr.AddReference("PresentationCore")
    clr.AddReference("WindowsBase")
    from System.Windows import LogicalTreeHelper, Visibility
    from System.Windows.Controls.Primitives import PlacementMode

    _all_schedule_views = collect_schedules(doc)
    all_schedule_names = sorted(v.Name for v in _all_schedule_views)
    all_schedules_by_name = {v.Name: v for v in _all_schedule_views}
    current_settings = dict(default_settings) if default_settings else {
        "export_mode": 0, "export_title": False, "export_column_headers": True,
        "export_group_headers": False, "export_grouped_column_headers": False,
        "use_category_sheet_name": True, "csv_mode": 0, "csv_delim": ",", "csv_text_qualifier": "",
    }

    paths = {}
    set_modes = {}
    name_boxes = {}
    schedule_combos = {}
    sheet_name_labels = {}
    path_labels = {}
    checkboxes = {}
    _ok_clicked = [False]

    def _snapshot():
        return json.dumps(saved_sets, sort_keys=True)

    _saved_snapshot = [_snapshot()]

    window = _load_dialog_xaml("BatchExportDialog.xaml")
    apply_window_title(window, "Export Schedules")
    window.Height = min(380 + max(len(saved_sets), 2) * 46, 720)

    rows_host = window.FindName("RowsHost")
    empty_state = window.FindName("EmptyState")
    set_count_text = window.FindName("SetCountText")
    row_template = window.FindResource("SetRowTemplate")
    text_brush = window.FindResource("TextBrush")
    muted_brush = window.FindResource("MutedBrush")
    try:
        (ui or load_uiutils()).uiUtils_load_logo(window.FindName("LogoImage"))
    except Exception:
        pass

    def _row_part(row, name):
        return LogicalTreeHelper.FindLogicalNode(row, name)

    def _set_path_label(label, path):
        shown = _display_path(path)
        label.Text = shown or "Not set"
        label.ToolTip = _expand_path(path or "") or None
        label.Foreground = text_brush if shown else muted_brush

    def _combo_value(combo, args=None):
        # While SelectionChanged runs, an editable ComboBox still holds the old Text,
        # so read the newly selected item from the event when there is one.
        try:
            if args is not None and args.AddedItems.Count:
                return str(args.AddedItems[0]).strip()
        except Exception:
            pass
        return str(combo.Text or (str(combo.SelectedItem) if combo.SelectedItem is not None else "")).strip()

    def _sheet_name_for(sel):
        # Blank until the text names a schedule in this model: a half-typed filter
        # is not a sheet name.
        view = all_schedules_by_name.get(sel) if sel else None
        if view is None:
            return ""
        if current_settings.get("use_category_sheet_name", True):
            return sanitize_sheet_name(_pluralize(_get_schedule_category_name(doc, view)))
        return sanitize_sheet_name(sel)

    def _collect_edits():
        """Fold the rows' edits into saved_sets; returns {old name: new name} for renames.
        Only real changes are written, so opening and closing the window changes nothing."""
        for sname, combo in list(schedule_combos.items()):
            sdata = saved_sets.get(sname)
            if sdata is None:
                continue
            sel = _combo_value(combo)
            current = sdata.get("schedule_names") or []
            # Older sets can list several schedules; keep them unless the row now names
            # another schedule in this model (a half-typed filter is not saved).
            if sel in all_schedules_by_name and sel != (current[0] if current else ""):
                sdata["schedule_names"] = [sel]
        for sname, path_val in list(paths.items()):
            sdata = saved_sets.get(sname)
            if sdata is None or not path_val:
                continue
            key = "excel_path" if set_modes.get(sname, 0) == 0 else "csv_folder"
            if path_val != _portable_path(_expand_path(sdata.get(key) or "")):
                sdata[key] = path_val
        renames = {}
        for orig, box in list(name_boxes.items()):
            new = (box.Text or "").strip()
            if new and new != orig and orig in saved_sets and new not in saved_sets:
                saved_sets[new] = saved_sets.pop(orig)
                if orig in paths:
                    paths[new] = paths.pop(orig)
                if orig in set_modes:
                    set_modes[new] = set_modes.pop(orig)
                renames[orig] = new
        return renames

    def _write_sets():
        proj_data_back = _normalize_namespace_data(read_saved_sets(doc, saved_sets_param))
        proj_data_back["sets"] = saved_sets
        ok_saved = write_saved_sets(doc, proj_data_back, saved_sets_param)
        if ok_saved:
            _saved_snapshot[0] = _snapshot()
        return ok_saved

    def _make_browse(sname_, mode_):
        def _on_browse(_s, _e):
            cur = _expand_path(paths.get(sname_) or "")
            if mode_ == 0:
                init_dir = os.path.dirname(cur) if cur else get_default_dir(doc)
                fname = os.path.basename(cur) if cur else "{}.xlsx".format(sanitize_file_name(sname_))
                new_path = _pick_save_file(
                    title="'{}' -- Choose Output File".format(sname_),
                    filter_text="Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                    default_extension="xlsx", initial_directory=init_dir, file_name=fname,
                )
                new_path = normalize_excel_output_path(new_path or "")
            else:
                init_dir = cur if cur and os.path.isdir(cur) else get_default_dir(doc)
                new_path = _pick_folder("'{}' -- Choose CSV Folder".format(sname_), initial_directory=init_dir)
                new_path = (new_path or "").strip()
            if new_path:
                paths[sname_] = _portable_path(new_path)
                _set_path_label(path_labels[sname_], paths[sname_])
        return _on_browse

    def _make_export_row(sname_):
        def _on_export(_s, _e):
            sdata = dict(saved_sets.get(sname_) or {})
            combo = schedule_combos.get(sname_)
            if combo:
                sel = _combo_value(combo)
                if sel:
                    sdata["schedule_names"] = [sel]
            mode = int(sdata.get("export_mode", current_settings.get("export_mode", 0)))
            path = _expand_path(paths.get(sname_) or "")
            if not path:
                init_dir = get_default_dir(doc)
                safe = re.sub(r'[\\/:*?"<>|]', "_", sname_)
                if mode == 0:
                    new_path = _pick_save_file("'{}' -- Output File".format(sname_),
                        "Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm", "xlsx", init_dir,
                        "{}.xlsx".format(safe))
                    new_path = normalize_excel_output_path(new_path or "")
                else:
                    new_path = _pick_folder("'{}' -- CSV Folder".format(sname_), init_dir)
                    new_path = (new_path or "").strip()
                if not new_path:
                    return
                paths[sname_] = _portable_path(new_path)
                if sname_ in path_labels:
                    _set_path_label(path_labels[sname_], paths[sname_])
                path = new_path
            all_sched = {v.Name: v for v in collect_schedules(doc)}
            sched_names = sdata.get("schedule_names") or []
            views = [all_sched[n] for n in sched_names if n in all_sched]
            if not views:
                if ui:
                    ui.uiUtils_alert("No schedule '{}' found in this model.".format(", ".join(sched_names) or "(none)"), title="Export")
                return
            cs = current_settings
            ok_val = None
            if mode == 0:
                locked = []
                ok_val = export_to_excel(doc, views, path, ui,
                    export_title=bool(sdata.get("export_title", cs.get("export_title", False))),
                    export_column_headers=bool(sdata.get("export_column_headers", cs.get("export_column_headers", True))),
                    export_group_headers=bool(sdata.get("export_group_headers", cs.get("export_group_headers", False))),
                    export_grouped_column_headers=bool(sdata.get("export_grouped_column_headers", cs.get("export_grouped_column_headers", False))),
                    text_qualifier=sdata.get("csv_text_qualifier", cs.get("csv_text_qualifier", "")),
                    delimiter=sdata.get("csv_delim", cs.get("csv_delim", ",")),
                    use_category_sheet_name=bool(cs.get("use_category_sheet_name", True)),
                    locked_messages=locked)
                if locked:
                    msg = locked[0]
                else:
                    msg = "Exported '{}' successfully.".format(sname_) if ok_val else "Export failed for '{}'. Check the log.".format(sname_)
            else:
                try:
                    export_to_csv(doc, views, path,
                        quote_all=bool(int(sdata.get("csv_mode", cs.get("csv_mode", 0))) == 1),
                        delimiter=sdata.get("csv_delim", cs.get("csv_delim", ",")),
                        export_title=bool(sdata.get("export_title", cs.get("export_title", False))),
                        export_column_headers=bool(sdata.get("export_column_headers", cs.get("export_column_headers", True))),
                        export_group_headers=bool(sdata.get("export_group_headers", cs.get("export_group_headers", False))),
                        export_grouped_column_headers=bool(sdata.get("export_grouped_column_headers", cs.get("export_grouped_column_headers", False))),
                        text_qualifier=sdata.get("csv_text_qualifier", cs.get("csv_text_qualifier", "")))
                    msg = "Exported '{}' to CSV.".format(sname_)
                except Exception as exc:
                    log_exception("inline export '{}'".format(sname_), exc)
                    msg = "Export failed for '{}'. Check the log.".format(sname_)
            if ui:
                ui.uiUtils_alert(msg, title="Batch Export")
        return _on_export

    def _make_delete_row(sname_):
        def _on_delete(_s, _e):
            name = _collect_edits().get(sname_, sname_)
            saved_sets.pop(name, None)
            paths.pop(name, None)
            _rebuild_rows()
        return _on_delete

    def _wire_combo(combo_, lbl_, names_):
        _busy = [False]
        def _update_lbl(args=None):
            lbl_.Text = _sheet_name_for(_combo_value(combo_, args))
        def _on_key_up(s, e):
            from System.Windows.Input import Key
            if _busy[0]:
                return
            if e.Key in (Key.Up, Key.Down, Key.Enter, Key.Return, Key.Tab, Key.Escape):
                _update_lbl()
                return
            text = str(combo_.Text or "").strip().lower()
            filtered = [n for n in names_ if text in n.lower()] if text else names_
            saved_text = combo_.Text
            _busy[0] = True
            combo_.Items.Clear()
            for n in filtered:
                combo_.Items.Add(n)
            combo_.Text = saved_text
            _busy[0] = False
            if filtered and not combo_.IsDropDownOpen:
                combo_.IsDropDownOpen = True
            _update_lbl()
        def _on_sel_changed(s, e):
            if not _busy[0]:
                _update_lbl(e)
        combo_.PreviewKeyUp += _on_key_up
        combo_.SelectionChanged += _on_sel_changed
        _update_lbl()

    def _rebuild_rows():
        rows_host.Children.Clear()
        checkboxes.clear(); path_labels.clear(); name_boxes.clear()
        schedule_combos.clear(); sheet_name_labels.clear()

        current_names = sorted(saved_sets.keys())
        set_count_text.Text = "1 set" if len(current_names) == 1 else "{} sets".format(len(current_names))
        empty_state.Visibility = Visibility.Collapsed if current_names else Visibility.Visible

        for sname in current_names:
            sdata = saved_sets.get(sname) or {}
            mode = int(sdata.get("export_mode", current_settings.get("export_mode", 0)))
            set_modes[sname] = mode
            if sname not in paths:
                raw = sdata.get("excel_path") if mode == 0 else sdata.get("csv_folder")
                paths[sname] = _portable_path(_expand_path(raw or ""))

            row = row_template.LoadContent()

            checkboxes[sname] = _row_part(row, "RowCheck")

            name_box = _row_part(row, "NameBox")
            name_box.Text = sname
            name_boxes[sname] = name_box

            sched_combo = _row_part(row, "ScheduleCombo")
            for n in all_schedule_names:
                sched_combo.Items.Add(n)
            existing = sdata.get("schedule_names") or []
            if existing:
                sched_combo.Text = existing[0]
            schedule_combos[sname] = sched_combo

            sheet_lbl = _row_part(row, "SheetText")
            sheet_lbl.Text = _sheet_name_for(existing[0] if existing else "")
            sheet_name_labels[sname] = sheet_lbl
            _wire_combo(sched_combo, sheet_lbl, all_schedule_names)

            path_tb = _row_part(row, "PathText")
            _set_path_label(path_tb, paths.get(sname))
            path_labels[sname] = path_tb

            browse_btn = _row_part(row, "BrowseButton")
            browse_btn.ToolTip = "Choose the output .xlsx file" if mode == 0 else "Choose the CSV folder"
            browse_btn.Click += _make_browse(sname, mode)
            _row_part(row, "ExportButton").Click += _make_export_row(sname)
            _row_part(row, "DeleteButton").Click += _make_delete_row(sname)

            rows_host.Children.Add(row)

    _rebuild_rows()

    def _open_settings(_s, _e):
        new_s = _show_export_settings_popup(current_settings, ui, owner=window)
        if new_s is not None:
            current_settings.clear()
            current_settings.update(new_s)
            for sname, combo in list(schedule_combos.items()):
                lbl = sheet_name_labels.get(sname)
                if lbl:
                    lbl.Text = _sheet_name_for(_combo_value(combo))

    sets_file_button = window.FindName("SetsFileButton")

    def _open_sets_menu(_s, _e):
        menu = sets_file_button.ContextMenu
        menu.PlacementTarget = sets_file_button
        menu.Placement = PlacementMode.Bottom
        menu.IsOpen = True

    def _save_file(_s, _e):
        if _collect_edits():
            _rebuild_rows()
        _export_sets_to_file(saved_sets, ui, tool_name="Export2Ex")

    def _load_file(_s, _e):
        imported = _import_sets_from_file(ui)
        if imported is not None:
            _collect_edits()
            # Imported sets replace same-named ones, so drop those rows' pending paths.
            for name in imported:
                paths.pop(name, None)
                set_modes.pop(name, None)
            saved_sets.update(imported)
            _rebuild_rows()

    def _add(_s, _e):
        _collect_edits()
        idx = 1
        while "New Set {}".format(idx) in saved_sets:
            idx += 1
        new_name = "New Set {}".format(idx)
        saved_sets[new_name] = dict(current_settings)
        saved_sets[new_name]["schedule_names"] = []
        _rebuild_rows()
        try:
            name_boxes[new_name].Focus()
            name_boxes[new_name].SelectAll()
        except Exception:
            pass

    def _ok(_s, _e):
        _ok_clicked[0] = True
        window.Close()

    # CancelButton ("Close") is IsCancel in the XAML: it closes the window (and so does Esc).
    window.FindName("SettingsButton").Click += _open_settings
    sets_file_button.Click += _open_sets_menu
    window.FindName("LoadFileMenuItem").Click += _load_file
    window.FindName("SaveFileMenuItem").Click += _save_file
    window.FindName("AddSetButton").Click += _add
    window.FindName("ExportSelectedButton").Click += _ok

    window.ShowDialog()

    checked = [sname for sname, box in list(checkboxes.items()) if box.IsChecked]
    renames = _collect_edits()
    if _snapshot() != _saved_snapshot[0] and not _write_sets() and ui:
        ui.uiUtils_alert(
            "Could not save the set changes to Project Information. "
            "In a workshared model, check that nobody else is editing Project Information.",
            title="Export Schedules")

    if not _ok_clicked[0]:
        return None

    selected = sorted(renames.get(sname, sname) for sname in checked)
    return selected, dict((k, _expand_path(v)) for k, v in paths.items()), dict(current_settings)


def read_csv_rows(path, delimiter=",", quotechar=""):
    log_message(
        "read_csv_rows start {} delimiter={!r} quotechar={!r}".format(
            describe_file_state(path), delimiter, quotechar
        )
    )
    for encoding in ("utf-8-sig", "utf-16", "cp1252"):
        for attempt in range(1, 4):
            try:
                text = read_text_file(path, encoding=encoding)
                handle = io.StringIO(text)
                if quotechar in ("\"", "'"):
                    reader = csv.reader(handle, delimiter=delimiter, quotechar=quotechar)
                else:
                    reader = csv.reader(handle, delimiter=delimiter)
                rows = [row for row in reader]
                log_message(
                    "read_csv_rows success path='{}' encoding={} attempt={} rows={}".format(
                        path, encoding, attempt, len(rows)
                    )
                )
                return rows
            except Exception as exc:
                log_message(
                    "read_csv_rows failed path='{}' encoding={} attempt={} {} error={}".format(
                        path, encoding, attempt, describe_file_state(path), str(exc)
                    )
                )
                if attempt < 3:
                    time.sleep(0.2)
    return []


def get_text_encoding(name):
    value = (name or "").strip().lower()
    if value == "utf-8-sig":
        return UTF8Encoding(True)
    if value == "utf-8":
        return UTF8Encoding(False)
    if value == "utf-16":
        return Encoding.Unicode
    return Encoding.GetEncoding(name)


def read_text_file(path, encoding="utf-8"):
    log_message("read_text_file open {} encoding={}".format(describe_file_state(path), encoding))
    reader = StreamReader(path, get_text_encoding(encoding), True)
    try:
        text = reader.ReadToEnd()
        log_message("read_text_file success path='{}' chars={}".format(path, len(text)))
        return text
    finally:
        reader.Close()


def write_text_file(path, text, encoding="utf-8"):
    log_message(
        "write_text_file start path='{}' encoding={} chars={} parent_exists={}".format(
            path,
            encoding,
            len(text or ""),
            os.path.isdir(os.path.dirname(path)) if os.path.dirname(path) else True,
        )
    )
    for attempt in range(1, 4):
        writer = None
        try:
            writer = StreamWriter(path, False, get_text_encoding(encoding))
            writer.Write(text)
            writer.Flush()
            log_message(
                "write_text_file success path='{}' attempt={} {}".format(
                    path, attempt, describe_file_state(path)
                )
            )
            return
        except Exception as exc:
            log_message(
                "write_text_file failed path='{}' attempt={} {} error={}".format(
                    path, attempt, describe_file_state(path), str(exc)
                )
            )
            if attempt >= 3:
                raise
            time.sleep(0.2)
        finally:
            if writer is not None:
                try:
                    writer.Close()
                except Exception:
                    pass


def normalize_table_data(data):
    return data


def get_section_data(view, section_type):
    table = view.GetTableData()
    try:
        return table.GetSectionData(section_type)
    except Exception:
        return None


def get_section_row_count(view, section_type):
    section = get_section_data(view, section_type)
    if section is None:
        return 0
    try:
        return section.NumberOfRows
    except Exception:
        return 0


def is_key_schedule(view):
    try:
        definition = view.Definition
    except Exception:
        definition = None
    try:
        return bool(definition and definition.IsKeySchedule)
    except Exception:
        return False


def get_body_row_element_ids(view):
    section = get_section_data(view, DB.SectionType.Body)
    if section is None:
        return []
    ids = []
    for row in range(section.NumberOfRows):
        elem_value = ""
        try:
            col_count = section.NumberOfColumns
        except Exception:
            col_count = 0
        for col in range(col_count):
            try:
                elem_id = section.GetCellElementId(row, col)
            except Exception:
                elem_id = None
            elem_int = element_id_value(elem_id)
            if elem_int != -1:
                elem_value = str(elem_int)
                break
        ids.append(elem_value)
    return ids


def inject_element_id_column(data, view, csv_text=False):
    if not data:
        return data
    header_rows = get_section_row_count(view, DB.SectionType.Header)
    body_ids = get_body_row_element_ids(view)
    total_rows = len(data)
    if header_rows > total_rows:
        header_rows = total_rows
    body_rows = min(len(body_ids), max(0, total_rows - header_rows))
    for idx, row in enumerate(data):
        if row is None:
            row = []
        if idx < header_rows:
            if idx == max(0, header_rows - 1):
                row.append("ElementId")
            else:
                row.append("")
        elif idx < header_rows + body_rows:
            elem_value = body_ids[idx - header_rows]
            if csv_text and elem_value and elem_value.isdigit() and len(elem_value) > 11:
                elem_value = "'" + elem_value
            row.append(elem_value)
        else:
            row.append("")
        data[idx] = row
    return data




def write_table_to_sheet(sheet, data, start_row, header_rows=0, column_specs=None, doc=None):
    if not data:
        return
    has_specs = bool(column_specs) and any(spec is not None for spec in column_specs)
    row_idx = start_row
    for row_offset, row in enumerate(data):
        col_idx = 1
        for value in row:
            spec = None
            if has_specs and column_specs and (col_idx - 1) < len(column_specs):
                spec = column_specs[col_idx - 1]
            if row_offset < header_rows:
                cell_value = value
            elif has_specs:
                if spec is None:
                    cell_value = value
                else:
                    cell_value = coerce_cell_value(value, spec=spec, doc=doc, numeric_fallback=True)
            else:
                cell_value = coerce_cell_value(value, spec=None, doc=None, numeric_fallback=False)
            cell = sheet.cell(row=row_idx, column=col_idx, value=cell_value)
            if row_offset >= header_rows and col_idx == len(row):
                if cell_value is None:
                    cell_value = ""
                cell.value = str(cell_value)
                cell.number_format = "@"
            col_idx += 1
        row_idx += 1


_NUMERIC_RE = re.compile(r"^-?\d+(\.\d+)?$")


def coerce_cell_value(value, spec=None, doc=None, numeric_fallback=True):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return ""
        if not numeric_fallback:
            return value
        if "'" in text or '"' in text or "/" in text:
            return value
        if re.search(r"\d[A-Za-z]", text):
            return value
        if _NUMERIC_RE.match(text):
            if spec is None:
                if text.startswith("0") and len(text) > 1 and not text.startswith("0."):
                    return value
                if text.startswith("-0") and len(text) > 2 and not text.startswith("-0."):
                    return value
            try:
                return int(text)
            except Exception:
                pass
            try:
                return float(text)
            except Exception:
                return value
    return value


def get_column_specs(view):
    section = get_section_data(view, DB.SectionType.Body)
    if section is None:
        return []
    try:
        col_count = int(section.NumberOfColumns)
    except Exception:
        col_count = 0
    specs = [None] * col_count
    try:
        definition = view.Definition
        field_ids = list(definition.GetFieldOrder())
        col = 0
        for field_id in field_ids:
            try:
                field = definition.GetField(field_id)
            except Exception:
                continue
            try:
                if field.IsHidden:
                    continue
            except Exception:
                pass
            if col >= col_count:
                break
            spec = None
            try:
                spec = field.GetSpecTypeId()
            except Exception:
                pass
            if spec is None:
                try:
                    spec = field.UnitType
                except Exception:
                    pass
            specs[col] = spec
            col += 1
    except Exception:
        pass
    return specs


def make_unique_name(base, used, max_len=None):
    candidate = base
    if max_len:
        candidate = candidate[:max_len]
    if candidate not in used:
        used.add(candidate)
        return candidate
    idx = 1
    while True:
        suffix = "_{}".format(idx)
        trunc = candidate
        if max_len:
            trunc = candidate[: max_len - len(suffix)]
        name = "{}{}".format(trunc, suffix)
        if name not in used:
            used.add(name)
            return name
        idx += 1


def _try_set_option(options, names, value):
    for name in names:
        if not hasattr(options, name):
            continue
        try:
            setattr(options, name, value)
            return True
        except Exception:
            continue
    return False


def apply_schedule_export_options(
    options,
    delimiter=",",
    export_title=False,
    export_column_headers=True,
    export_group_headers=False,
    export_grouped_column_headers=False,
    text_qualifier="",
):
    options.FieldDelimiter = delimiter
    # Revit API property names differ by version; support both.
    _try_set_option(options, ("ExportTitle", "Title"), bool(export_title))
    _try_set_option(options, ("ExportColumnHeaders", "ColumnHeaders"), bool(export_column_headers))
    _try_set_option(options, ("ExportGroupHeaders", "HeadersFootersBlanks"), bool(export_group_headers))
    _try_set_option(options, ("ExportGroupedColumnHeaders",), bool(export_grouped_column_headers))

    if text_qualifier in ("\"", "'"):
        _try_set_option(options, ("TextQualifier",), text_qualifier)
    else:
        _try_set_option(options, ("TextQualifier",), "")


def export_to_excel(
    doc,
    schedules,
    file_path,
    ui,
    export_title=False,
    export_column_headers=True,
    export_group_headers=False,
    export_grouped_column_headers=False,
    text_qualifier="",
    delimiter=",",
    use_category_sheet_name=False,
    locked_messages=None,
):
    """Export schedules into one workbook. Returns True on success. If the workbook is
    open in Excel, the "close it" message is appended to locked_messages when a list is
    given (batch export), otherwise shown as an alert, and False is returned."""
    log_message(
        "export_to_excel start file_path='{}' schedules={} delimiter={!r} quotechar={!r} title={} col_headers={} group_headers={} grouped_col_headers={}".format(
            file_path,
            len(schedules),
            delimiter,
            text_qualifier,
            bool(export_title),
            bool(export_column_headers),
            bool(export_group_headers),
            bool(export_grouped_column_headers),
        )
    )
    add_lib_path()
    try:
        import WWP_xlsx as openpyxl
    except Exception as exc:
        ui.uiUtils_alert(
            "Excel writer is not available.\n{}".format(exc),
            title="Multiple Schedules Exporter",
        )
        return False

    def _report_locked():
        message = openpyxl.file_locked_message(file_path)
        log_message("export_to_excel locked file_path='{}'".format(file_path))
        if locked_messages is not None:
            locked_messages.append(message)
        else:
            ui.uiUtils_alert(message, title="Multiple Schedules Exporter")
        return False

    if openpyxl.is_file_locked(file_path):
        return _report_locked()

    used_names = set()
    try:
        if os.path.exists(file_path):
            load_kwargs = {}
            if os.path.splitext(file_path)[1].lower() == ".xlsm":
                load_kwargs["keep_vba"] = True
            workbook = openpyxl.load_workbook(file_path, **load_kwargs)
        else:
            workbook = openpyxl.Workbook()
    except Exception as exc:
        if openpyxl.is_file_lock_error(exc):
            return _report_locked()
        raise
    temp_dir = tempfile.mkdtemp(prefix="wwp_schedules_")

    options = DB.ViewScheduleExportOptions()
    apply_schedule_export_options(
        options,
        delimiter=delimiter,
        export_title=export_title,
        export_column_headers=export_column_headers,
        export_group_headers=export_group_headers,
        export_grouped_column_headers=export_grouped_column_headers,
        text_qualifier=text_qualifier,
    )
    try:
        for view in schedules:
            log_message("export_to_excel schedule start name='{}' id={}".format(view.Name, element_id_value(view.Id)))
            key_schedule = is_key_schedule(view)
            raw_name = _pluralize(_get_schedule_category_name(doc, view)) if use_category_sheet_name else view.Name
            base_name = sanitize_sheet_name(raw_name)
            if base_name in workbook.sheetnames and base_name not in used_names:
                sheet_name = base_name
            else:
                used_pool = set(workbook.sheetnames)
                used_pool.update(used_names)
                sheet_name = make_unique_name(base_name, used_pool, max_len=31)
            used_names.add(sheet_name)
            if sheet_name in workbook.sheetnames:
                existing = workbook[sheet_name]
                sheet_idx = workbook.worksheets.index(existing)
                workbook.remove(existing)
                sheet = workbook.create_sheet(title=sheet_name, index=sheet_idx)
            else:
                sheet = workbook.create_sheet(title=sheet_name)

            temp_name = "{}.csv".format(sanitize_file_name(view.Name))
            log_message("export_to_excel view.Export temp_dir='{}' temp_name='{}'".format(temp_dir, temp_name))
            view.Export(temp_dir, temp_name, options)
            csv_path = os.path.join(temp_dir, temp_name)
            log_message("export_to_excel exported temp csv {}".format(describe_file_state(csv_path)))
            data = normalize_table_data(read_csv_rows(csv_path, delimiter=delimiter, quotechar=text_qualifier))
            if not key_schedule:
                data = inject_element_id_column(data, view, csv_text=False)
            header_rows = get_section_row_count(view, DB.SectionType.Header)
            column_specs = get_column_specs(view)
            if key_schedule:
                # Preserve key schedule values exactly as exported (avoid numeric coercion).
                column_specs = None
            elif column_specs is not None:
                column_specs = [None] + column_specs
            write_table_to_sheet(
                sheet,
                data,
                1,
                header_rows=header_rows,
                column_specs=column_specs,
                doc=doc,
            )
            log_message("export_to_excel schedule complete name='{}' rows={}".format(view.Name, len(data)))
    finally:
        log_message("export_to_excel cleanup temp_dir='{}'".format(temp_dir))
        shutil.rmtree(temp_dir, ignore_errors=True)

    if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
        default_sheet = workbook["Sheet"]
        workbook.remove(default_sheet)

    try:
        workbook.save(file_path)
    except Exception as exc:
        # Excel can open the file while the schedules are being exported.
        if openpyxl.is_file_lock_error(exc):
            return _report_locked()
        raise
    log_message("export_to_excel saved workbook {}".format(describe_file_state(file_path)))
    return True


def export_to_csv(
    doc,
    schedules,
    folder,
    quote_all=False,
    delimiter=",",
    export_title=False,
    export_column_headers=True,
    export_group_headers=False,
    export_grouped_column_headers=False,
    text_qualifier="",
):
    log_message(
        "export_to_csv start folder='{}' schedules={} delimiter={!r} quotechar={!r} quote_all={} title={} col_headers={} group_headers={} grouped_col_headers={}".format(
            folder,
            len(schedules),
            delimiter,
            text_qualifier,
            bool(quote_all),
            bool(export_title),
            bool(export_column_headers),
            bool(export_group_headers),
            bool(export_grouped_column_headers),
        )
    )
    if not os.path.isdir(folder):
        os.makedirs(folder)
        log_message("export_to_csv created folder='{}'".format(folder))
    options = DB.ViewScheduleExportOptions()
    apply_schedule_export_options(
        options,
        delimiter=delimiter,
        export_title=export_title,
        export_column_headers=export_column_headers,
        export_group_headers=export_group_headers,
        export_grouped_column_headers=export_grouped_column_headers,
        text_qualifier=text_qualifier,
    )
    used_names = set()
    temp_dir = tempfile.mkdtemp(prefix="wwp_schedules_csv_")
    log_message("export_to_csv temp_dir='{}'".format(temp_dir))
    try:
        for view in schedules:
            log_message("export_to_csv schedule start name='{}' id={}".format(view.Name, element_id_value(view.Id)))
            key_schedule = is_key_schedule(view)
            base_name = sanitize_file_name(view.Name)
            unique_name = make_unique_name(base_name, used_names)
            file_name = "{}.csv".format(unique_name)
            temp_name = "tmp_{}_{}.csv".format(element_id_value(view.Id), unique_name)
            temp_csv_path = os.path.join(temp_dir, temp_name)
            final_csv_path = os.path.join(folder, file_name)
            log_message("export_to_csv view.Export temp_dir='{}' temp_name='{}'".format(temp_dir, temp_name))
            view.Export(temp_dir, temp_name, options)
            log_message("export_to_csv exported temp csv {}".format(describe_file_state(temp_csv_path)))
            rows = normalize_table_data(read_csv_rows(temp_csv_path, delimiter=delimiter, quotechar=text_qualifier))
            if not key_schedule:
                rows = inject_element_id_column(rows, view, csv_text=True)
            buffer_handle = io.StringIO()
            if text_qualifier in ("\"", "'"):
                writer = csv.writer(
                    buffer_handle,
                    delimiter=delimiter,
                    quoting=csv.QUOTE_ALL if quote_all else csv.QUOTE_MINIMAL,
                    quotechar=text_qualifier,
                )
            else:
                writer = csv.writer(
                    buffer_handle,
                    delimiter=delimiter,
                    quoting=csv.QUOTE_ALL if quote_all else csv.QUOTE_MINIMAL,
                )
            writer.writerows(rows)
            write_text_file(final_csv_path, buffer_handle.getvalue(), encoding="utf-8-sig")
            log_message(
                "export_to_csv schedule complete name='{}' temp={} output={}".format(
                    view.Name,
                    describe_file_state(temp_csv_path),
                    describe_file_state(final_csv_path),
                )
            )
    finally:
        log_message("export_to_csv cleanup temp_dir='{}'".format(temp_dir))
        shutil.rmtree(temp_dir, ignore_errors=True)
    return True


def select_csv_mode(ui, default_mode=0):
    options = [
        "Standard CSV (comma, minimal quotes)",
        "CSV (Quoted - all fields)",
    ]
    try:
        selected = ui.uiUtils_select_indices(
            options,
            title="CSV Export Mode",
            prompt="Choose CSV export format:",
            multiselect=False,
            width=520,
            height=260,
        )
    except Exception:
        selected = []
    if selected is None or len(selected) == 0:
        return None
    if selected[0] < 0 or selected[0] >= len(options):
        return default_mode
    return int(selected[0])


def select_csv_delimiter(ui, default_delimiter=","):
    options = [
        ("Comma (,)", ","),
        ("Semicolon (;)", ";"),
        ("Tab (\\t)", "\t"),
    ]
    labels = [opt[0] for opt in options]
    default_index = 0
    for idx, opt in enumerate(options):
        if opt[1] == default_delimiter:
            default_index = idx
            break
    try:
        selected = ui.uiUtils_select_indices(
            labels,
            title="CSV Delimiter",
            prompt="Choose delimiter:",
            multiselect=False,
            width=520,
            height=260,
        )
    except Exception:
        selected = []
    if selected is None or len(selected) == 0:
        return None
    sel_idx = int(selected[0]) if selected else default_index
    if sel_idx < 0 or sel_idx >= len(options):
        sel_idx = default_index
    return options[sel_idx][1]


def main():
    log_message("main start log_path='{}'".format(_log_file_path()))
    doc = get_active_doc()
    if doc is None:
        ui = load_uiutils()
        ui.uiUtils_alert("No active Revit document found.", title="Multiple Schedules Exporter")
        return
    config, save_config = get_config_and_saver()
    ui = load_uiutils()
    saved_sets_param = _ensure_saved_sets_param(doc, config, ui, save_config)
    proj_data = _normalize_namespace_data(read_saved_sets(doc, saved_sets_param))
    project_settings = proj_data.get("settings", {}) if isinstance(proj_data, dict) else {}
    if not isinstance(project_settings, dict):
        project_settings = {}

    schedules = collect_schedules(doc)
    if not schedules:
        ui.uiUtils_alert("No schedules found.", title="Multiple Schedules Exporter")
        return

    items = [ScheduleItem(v) for v in schedules]
    log_message("main schedules loaded count={}".format(len(items)))
    last_ids = project_settings.get(CONFIG_LAST_SCHEDULE_IDS, config_get(config, CONFIG_LAST_SCHEDULE_IDS, []))
    try:
        prechecked_ids = set(int(x) for x in last_ids)
    except Exception:
        prechecked_ids = set()
    prechecked_indices = [
        idx for idx, item in enumerate(items)
        if element_id_value(item.view.Id) in prechecked_ids
    ]
    default_dir = get_default_dir(doc)
    last_excel_path = config_get(config, CONFIG_LAST_EXCEL_PATH, "")
    last_csv_dir = config_get(config, CONFIG_LAST_CSV_DIR, "")
    last_csv_mode = project_settings.get(CONFIG_LAST_CSV_MODE, config_get(config, CONFIG_LAST_CSV_MODE, 0))
    last_csv_delim = project_settings.get(CONFIG_LAST_CSV_DELIM, config_get(config, CONFIG_LAST_CSV_DELIM, ","))
    last_export_mode = project_settings.get(CONFIG_LAST_EXPORT_MODE, config_get(config, CONFIG_LAST_EXPORT_MODE, 0))
    last_export_title = project_settings.get(CONFIG_LAST_CSV_EXPORT_TITLE, config_get(config, CONFIG_LAST_CSV_EXPORT_TITLE, False))
    last_column_headers = project_settings.get(CONFIG_LAST_CSV_COLUMN_HEADERS, config_get(config, CONFIG_LAST_CSV_COLUMN_HEADERS, True))
    last_group_headers = project_settings.get(CONFIG_LAST_CSV_GROUP_HEADERS, config_get(config, CONFIG_LAST_CSV_GROUP_HEADERS, False))
    last_grouped_column_headers = project_settings.get(CONFIG_LAST_CSV_GROUPED_COLUMN_HEADERS, config_get(config, CONFIG_LAST_CSV_GROUPED_COLUMN_HEADERS, False))
    last_text_qualifier = project_settings.get(CONFIG_LAST_CSV_TEXT_QUALIFIER, config_get(config, CONFIG_LAST_CSV_TEXT_QUALIFIER, ""))
    last_use_category_sheet_name = project_settings.get(CONFIG_LAST_USE_CATEGORY_SHEET_NAME, config_get(config, CONFIG_LAST_USE_CATEGORY_SHEET_NAME, False))
    saved_sets = proj_data.get("sets", {})
    if not isinstance(saved_sets, dict):
        saved_sets = {}

    init_excel_path = last_excel_path or os.path.join(default_dir, "Schedules.xlsx")
    init_csv_dir = ensure_existing_dir(last_csv_dir, default_dir)

    default_settings = {
        "export_mode": last_export_mode,
        "export_title": bool(last_export_title),
        "export_column_headers": bool(last_column_headers),
        "export_group_headers": bool(last_group_headers),
        "export_grouped_column_headers": bool(last_grouped_column_headers),
        "use_category_sheet_name": True,
        "csv_mode": last_csv_mode,
        "csv_delim": last_csv_delim,
        "csv_text_qualifier": last_text_qualifier,
    }

    batch_result = _show_batch_dialog(
        saved_sets, doc,
        ui=ui,
        saved_sets_param=saved_sets_param,
        default_settings=default_settings,
    )
    if batch_result is None:
        return
    selected_batch_names, batch_override_paths, batch_settings = batch_result
    if not selected_batch_names:
        ui.uiUtils_alert("Select at least one set to export.", title="Multiple Schedules Exporter")
        return
    inputs = {"batch_sets": selected_batch_names, "batch_paths": batch_override_paths}
    if inputs is not False:
        if not inputs:
            return
        if "batch_sets" in inputs:
            batch_set_names = inputs["batch_sets"]
            batch_paths = inputs.get("batch_paths") or {}
            all_schedules_by_name = {v.Name: v for v in schedules}
            success_count = 0
            skipped = []
            warnings = []
            for set_name in batch_set_names:
                set_data = saved_sets.get(set_name)
                if not isinstance(set_data, dict):
                    skipped.append("{}: set not found".format(set_name))
                    continue
                set_mode = int(set_data.get("export_mode", 0))
                set_sched_names = set_data.get("schedule_names") or []
                set_views = [all_schedules_by_name[n] for n in set_sched_names if n in all_schedules_by_name]
                if not set_views:
                    skipped.append("{}: no matching schedules found in this model".format(set_name))
                    continue
                missing = [n for n in set_sched_names if n not in all_schedules_by_name]
                if missing:
                    sample = ", ".join(missing[:3]) + ("..." if len(missing) > 3 else "")
                    warnings.append("{}: {}/{} schedules not found (renamed?) -- {}".format(
                        set_name, len(set_views), len(set_sched_names), sample))
                set_use_cat = bool(batch_settings.get("use_category_sheet_name", True))
                bs = batch_settings
                set_exp_title = bool(set_data.get("export_title", bs.get("export_title", False)))
                set_col_hdr = bool(set_data.get("export_column_headers", bs.get("export_column_headers", True)))
                set_grp_hdr = bool(set_data.get("export_group_headers", bs.get("export_group_headers", False)))
                set_grpd_col = bool(set_data.get("export_grouped_column_headers", bs.get("export_grouped_column_headers", False)))
                set_delim = set_data.get("csv_delim", bs.get("csv_delim", ","))
                set_qualifier = set_data.get("csv_text_qualifier", bs.get("csv_text_qualifier", ""))
                if set_mode == 0:
                    set_excel_path = normalize_excel_output_path(batch_paths.get(set_name) or "")
                    if not set_excel_path:
                        skipped.append("{}: no output path set".format(set_name))
                        continue
                    locked_messages = []
                    ok = export_to_excel(
                        doc, set_views, set_excel_path, ui,
                        export_title=set_exp_title,
                        export_column_headers=set_col_hdr,
                        export_group_headers=set_grp_hdr,
                        export_grouped_column_headers=set_grpd_col,
                        text_qualifier=set_qualifier,
                        delimiter=set_delim,
                        use_category_sheet_name=set_use_cat,
                        locked_messages=locked_messages,
                    )
                    if locked_messages:
                        skipped.append("{}: {}".format(set_name, locked_messages[0]))
                    elif ok:
                        success_count += 1
                    else:
                        skipped.append("{}: export failed (see log)".format(set_name))
                else:
                    set_csv_folder = (batch_paths.get(set_name) or "").strip()
                    if not set_csv_folder:
                        skipped.append("{}: no output folder set".format(set_name))
                        continue
                    try:
                        export_to_csv(
                            doc, set_views, set_csv_folder,
                            quote_all=bool(int(set_data.get("csv_mode", 0)) == 1),
                            delimiter=set_delim,
                            export_title=set_exp_title,
                            export_column_headers=set_col_hdr,
                            export_group_headers=set_grp_hdr,
                            export_grouped_column_headers=set_grpd_col,
                            text_qualifier=set_qualifier,
                        )
                        success_count += 1
                    except Exception as _csv_exc:
                        log_exception("batch CSV export set '{}'".format(set_name), _csv_exc)
                        skipped.append("{}: export failed (see log)".format(set_name))
            paths_changed = False
            for sname, new_path in batch_paths.items():
                sdata = saved_sets.get(sname)
                if not isinstance(sdata, dict):
                    continue
                mode = int(sdata.get("export_mode", 0))
                if mode == 0:
                    new_path = _portable_path(new_path)
                    if new_path and new_path != (sdata.get("excel_path") or "").strip():
                        sdata["excel_path"] = new_path
                        paths_changed = True
                else:
                    new_path = _portable_path(new_path)
                    if new_path and new_path != (sdata.get("csv_folder") or "").strip():
                        sdata["csv_folder"] = new_path
                        paths_changed = True
            if paths_changed:
                proj_data_back = _normalize_namespace_data(read_saved_sets(doc, saved_sets_param))
                proj_data_back["sets"] = saved_sets
                write_saved_sets(doc, proj_data_back, saved_sets_param)

            msg = "{} of {} set(s) exported successfully.".format(success_count, len(batch_set_names))
            if warnings:
                msg += "\n\nPartial exports (some schedules not found):\n" + "\n".join(warnings)
            if skipped:
                msg += "\n\nSkipped:\n" + "\n".join(skipped)
            ui.uiUtils_alert(msg, title="Multiple Schedules Exporter")
            return

        selected_indices = inputs.get("selected_indices") or []
        if not selected_indices:
            ui.uiUtils_alert("Select at least one schedule.", title="Multiple Schedules Exporter")
            return
        selected_views = [items[i].view for i in selected_indices]
        config.last_schedule_ids = [element_id_value(v.Id) for v in selected_views]
        export_title = bool(inputs.get("export_title"))
        export_column_headers = bool(inputs.get("export_column_headers"))
        export_group_headers = bool(inputs.get("export_group_headers"))
        export_grouped_column_headers = bool(inputs.get("export_grouped_column_headers"))
        csv_text_qualifier = inputs.get("csv_text_qualifier") or ""
        csv_delim = inputs.get("csv_delimiter") or last_csv_delim
        quote_all = bool(inputs.get("csv_quote_all"))
        use_category = bool(inputs.get("use_category_sheet_name", False))
        mode = int(inputs.get("mode", 0))
        if mode == 0:
            file_path = normalize_excel_output_path(inputs.get("excel_path"))
            if not file_path:
                ui.uiUtils_alert(
                    "Choose an Excel file path ending with .xlsx or .xlsm.",
                    title="Multiple Schedules Exporter",
                )
                return
            success = export_to_excel(
                doc,
                selected_views,
                file_path,
                ui,
                export_title=export_title,
                export_column_headers=export_column_headers,
                export_group_headers=export_group_headers,
                export_grouped_column_headers=export_grouped_column_headers,
                text_qualifier=csv_text_qualifier,
                delimiter=csv_delim,
                use_category_sheet_name=use_category,
            )
            if not success:
                return
            config.last_excel_path = _normalize_path(file_path)
            try:
                os.startfile(file_path)
            except Exception:
                pass
        else:
            folder = os.path.expandvars((inputs.get("csv_folder") or "").strip())
            if not folder:
                ui.uiUtils_alert("Choose a CSV folder.", title="Multiple Schedules Exporter")
                return
            export_to_csv(
                doc,
                selected_views,
                folder,
                quote_all=quote_all,
                delimiter=csv_delim,
                export_title=export_title,
                export_column_headers=export_column_headers,
                export_group_headers=export_group_headers,
                export_grouped_column_headers=export_grouped_column_headers,
                text_qualifier=csv_text_qualifier,
            )
            config.last_csv_dir = _normalize_path(folder)
        config.last_export_mode = mode
        config.last_csv_mode = 1 if quote_all else 0
        config.last_csv_delim = csv_delim
        config.last_csv_text_qualifier = csv_text_qualifier
        config.last_csv_export_title = export_title
        config.last_csv_column_headers = export_column_headers
        config.last_csv_group_headers = export_group_headers
        config.last_csv_grouped_column_headers = export_grouped_column_headers
        config.last_use_category_sheet_name = use_category
        proj_data = _normalize_namespace_data(read_saved_sets(doc, saved_sets_param))
        proj_data["settings"] = {
            CONFIG_LAST_SCHEDULE_IDS: [element_id_value(v.Id) for v in selected_views],
            CONFIG_LAST_EXPORT_MODE: mode,
            CONFIG_LAST_CSV_MODE: 1 if quote_all else 0,
            CONFIG_LAST_CSV_DELIM: csv_delim,
            CONFIG_LAST_CSV_EXPORT_TITLE: export_title,
            CONFIG_LAST_CSV_COLUMN_HEADERS: export_column_headers,
            CONFIG_LAST_CSV_GROUP_HEADERS: export_group_headers,
            CONFIG_LAST_CSV_GROUPED_COLUMN_HEADERS: export_grouped_column_headers,
            CONFIG_LAST_CSV_TEXT_QUALIFIER: csv_text_qualifier,
            CONFIG_LAST_USE_CATEGORY_SHEET_NAME: use_category,
        }
        proj_data["sets"] = proj_data.get("sets", {})
        write_saved_sets(doc, proj_data, saved_sets_param)
        log_message("main saving config after modern dialog")
        save_config()
        ui.uiUtils_alert("Export complete.", title="Multiple Schedules Exporter")
        return

    if hasattr(ui, "uiUtils_select_items_with_mode"):
        selected_indices, mode = ui.uiUtils_select_items_with_mode(
            [item.display_name for item in items],
            title="Export Schedules",
            prompt="Select schedules to export:",
            mode_labels=("Export to Excel", "Export to CSV"),
            default_mode=0,
            prechecked_indices=prechecked_indices,
            width=680,
            height=620,
        )
    else:
        ui.uiUtils_alert(
            "UI helper uiUtils_select_items_with_mode is unavailable. Restart pyRevit or update WWP_uiUtils.",
            title="Multiple Schedules Exporter",
        )
        return
    if mode is None:
        return
    if not selected_indices:
        ui.uiUtils_alert("Select at least one schedule.", title="Multiple Schedules Exporter")
        return
    selected_views = [items[i].view for i in selected_indices]
    config.last_schedule_ids = [element_id_value(v.Id) for v in selected_views]

    if mode == 0:
        last_excel_dir = os.path.dirname(last_excel_path) if last_excel_path else ""
        init_dir = ensure_existing_dir(last_excel_dir, default_dir)
        file_name = os.path.basename(last_excel_path) if last_excel_path else "Schedules.xlsx"
        try:
            file_path = _pick_save_file(
                title="Export Schedules",
                filter_text="Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                default_extension="xlsx",
                initial_directory=init_dir,
                file_name=file_name,
            )
        except Exception as exc:
            log_exception("Legacy native save dialog failed", exc)
            try:
                file_path = _pick_save_file(
                    title="Export Schedules",
                    filter_text="Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                    default_extension="xlsx",
                    initial_directory="",
                    file_name=file_name,
                )
            except Exception as retry_exc:
                log_exception("Legacy native save dialog retry failed", retry_exc)
                ui.uiUtils_alert(
                    "Could not open the Excel save dialog. Check the suggested path and try again.",
                    title="Multiple Schedules Exporter",
                )
                return
        file_path = normalize_excel_output_path(file_path)
        if not file_path:
            ui.uiUtils_alert(
                "Choose an Excel file path ending with .xlsx or .xlsm.",
                title="Multiple Schedules Exporter",
            )
            return
        success = export_to_excel(doc, selected_views, file_path, ui)
        if not success:
            return
        config.last_excel_path = _normalize_path(file_path)
        try:
            os.startfile(file_path)
        except Exception:
            pass
    else:
        init_dir = ensure_existing_dir(last_csv_dir, default_dir)
        csv_mode = select_csv_mode(ui, default_mode=last_csv_mode)
        if csv_mode is None:
            return
        csv_delim = select_csv_delimiter(ui, default_delimiter=last_csv_delim)
        if csv_delim is None:
            return
        folder = _pick_folder(
            title="Select CSV Folder",
            initial_directory=init_dir,
        )
        if not folder:
            return
        export_to_csv(doc, selected_views, folder, quote_all=(csv_mode == 1), delimiter=csv_delim)
        config.last_csv_dir = _normalize_path(folder)
        config.last_csv_mode = csv_mode
        config.last_csv_delim = csv_delim

    log_message("main saving config after legacy dialog")
    proj_data = _normalize_namespace_data(read_saved_sets(doc, saved_sets_param))
    proj_data["settings"] = {
        CONFIG_LAST_SCHEDULE_IDS: [element_id_value(v.Id) for v in selected_views],
        CONFIG_LAST_EXPORT_MODE: mode,
        CONFIG_LAST_CSV_MODE: config_get(config, CONFIG_LAST_CSV_MODE, 0),
        CONFIG_LAST_CSV_DELIM: config_get(config, CONFIG_LAST_CSV_DELIM, ","),
        CONFIG_LAST_CSV_EXPORT_TITLE: config_get(config, CONFIG_LAST_CSV_EXPORT_TITLE, False),
        CONFIG_LAST_CSV_COLUMN_HEADERS: config_get(config, CONFIG_LAST_CSV_COLUMN_HEADERS, True),
        CONFIG_LAST_CSV_GROUP_HEADERS: config_get(config, CONFIG_LAST_CSV_GROUP_HEADERS, False),
        CONFIG_LAST_CSV_GROUPED_COLUMN_HEADERS: config_get(config, CONFIG_LAST_CSV_GROUPED_COLUMN_HEADERS, False),
        CONFIG_LAST_CSV_TEXT_QUALIFIER: config_get(config, CONFIG_LAST_CSV_TEXT_QUALIFIER, ""),
    }
    proj_data["sets"] = proj_data.get("sets", {})
    write_saved_sets(doc, proj_data, saved_sets_param)
    save_config()
    ui.uiUtils_alert("Export complete.", title="Multiple Schedules Exporter")


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        log_exception("Unhandled exception in Export2Ex", ex)
        show_error_report(ex)
        raise
