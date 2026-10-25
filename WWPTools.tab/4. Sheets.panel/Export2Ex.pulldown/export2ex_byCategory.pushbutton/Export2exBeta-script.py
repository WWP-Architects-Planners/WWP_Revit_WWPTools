# -*- coding: utf-8 -*-
import json
import os
import re
import sys
import time
import traceback

import clr
clr.AddReference('System.Xml')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
from System import String, Int64
from System import Object
from System.Collections.Generic import List
from System.IO import File
from System.Windows.Media import Brushes
from System.Windows.Controls import ListBoxItem

from pyrevit import DB
from WWP_settings import get_tool_settings
from WWP_versioning import apply_window_title


CONFIG_LAST_EXCEL_PATH = "last_excel_path"
CONFIG_LAST_MODE = "last_mode"
CONFIG_LAST_SCHEDULE_ID = "last_schedule_id"
CONFIG_LAST_CATEGORY_ID = "last_category_id"
CONFIG_LAST_PARAM_NAMES = "last_param_names"
CONFIG_LAST_SHEET_NAME = "last_sheet_name"
PARAM_SAVED_SETS = "! P_STATS_Export_Text"
SAVED_SET_NAMESPACE = "export2ex_beta"
LOG_FILE_NAME = "Export2ExBeta.log"
ALLOWED_EXCEL_EXTENSIONS = (".xlsx", ".xlsm")
_last_write_error = ""
MODE_FROM_SCHEDULE = "schedule"

class _JsonEncoder(json.JSONEncoder):
    def default(self, obj):
        if type(obj).__name__ in ('long', 'Int64', 'Int32'):
            return int(obj)
        return super(_JsonEncoder, self).default(obj)
MODE_BY_CATEGORY = "category"
EMBEDDED_EXPORT_DIALOG_XAML = r'''<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Export2Ex - By Category"
        Height="780"
        Width="1160"
        MinHeight="620"
        MinWidth="940"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResizeWithGrip"
        Background="{StaticResource WindowBrush}">
    <Grid Margin="{StaticResource PageMargin}">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,16">
            <TextBlock Text="Export set" Style="{StaticResource PageTitle}" />
            <TextBlock Style="{StaticResource PageSubtitle}"
                       Text="Pick a category (or a schedule, to use its category and fields), then choose the parameters to export in column order." />
        </StackPanel>

        <!-- Set details -->
        <Border Grid.Row="1" Style="{StaticResource WwpCard}" Padding="20,16">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="16" />
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="16" />
                    <ColumnDefinition Width="200" />
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Set name" Style="{StaticResource FieldLabel}" />
                    <TextBox Name="SetNameBox" AutomationProperties.Name="Set name" />
                </StackPanel>
                <StackPanel Grid.Column="2">
                    <TextBlock Text="Sheet name" Style="{StaticResource FieldLabel}" />
                    <TextBox Name="SheetNameBox" AutomationProperties.Name="Sheet name" />
                </StackPanel>
                <StackPanel Grid.Column="4">
                    <TextBlock Text="Units" Style="{StaticResource FieldLabel}" />
                    <ComboBox Name="UnitsBox" SelectedIndex="0" AutomationProperties.Name="Units">
                        <ComboBoxItem Content="Project units" />
                        <ComboBoxItem Content="Imperial" />
                        <ComboBoxItem Content="Metric" />
                    </ComboBox>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Source, available parameters, exported columns -->
        <Border Grid.Row="2" Style="{StaticResource WwpCard}" Margin="0,12,0,0" Padding="16">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="300" />
                    <ColumnDefinition Width="14" />
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="150" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>

                <!-- SOURCE -->
                <Border Grid.Column="0" Style="{StaticResource TableLane}" Margin="0"
                        Background="{StaticResource LaneGreyBrush}" Padding="12,0,12,12">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="SOURCE" Style="{StaticResource TableGroupLabel}"
                                   Margin="0,10,0,10" />
                        <CheckBox Grid.Row="1" Name="FromScheduleToggle"
                                  Style="{StaticResource ToggleSwitchCheckBox}"
                                  Content="Category from schedule"
                                  ToolTip="Pick a schedule instead: its category and visible fields are used"
                                  Margin="0,0,0,12" />
                        <TextBlock Grid.Row="2" Name="SourceLabel" Text="Search categories"
                                   Style="{StaticResource FieldLabel}" />
                        <TextBox Grid.Row="3" Name="SourceSearchBox" Margin="0,0,0,10"
                                 AutomationProperties.Name="Search sources" />
                        <TextBlock Grid.Row="4" Name="SourceListLabel" Text="Categories"
                                   Style="{StaticResource FieldLabel}" />
                        <ListBox Grid.Row="5" Name="SourceList" SelectionMode="Single" MinHeight="120"
                                 AutomationProperties.Name="Sources" />
                    </Grid>
                </Border>

                <!-- AVAILABLE -->
                <Border Grid.Column="2" Style="{StaticResource TableLane}" Margin="0"
                        Background="{StaticResource LaneBlueBrush}" Padding="12,0,12,12">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="AVAILABLE PARAMETERS" Style="{StaticResource TableGroupLabel}"
                                   Foreground="{StaticResource LaneBlueTextBrush}" Margin="0,10,0,10" />
                        <Grid Grid.Row="1" Margin="0,0,0,10">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*" />
                                <ColumnDefinition Width="Auto" />
                            </Grid.ColumnDefinitions>
                            <TextBox Name="ParameterSearchBox" Margin="0,0,12,0"
                                     ToolTip="Filter parameters by name"
                                     AutomationProperties.Name="Search parameters" />
                            <CheckBox Grid.Column="1" Name="ShowReadOnlyFilter" Content="Show read-only"
                                      IsChecked="True" />
                        </Grid>
                        <ListBox Grid.Row="2" Name="ParameterList" SelectionMode="Extended" MinHeight="120"
                                 AutomationProperties.Name="Available parameters" />
                    </Grid>
                </Border>

                <!-- Move buttons -->
                <StackPanel Grid.Column="3" VerticalAlignment="Center" Margin="12,0">
                    <Button Name="AddParameterButton" Margin="0,0,0,8" HorizontalContentAlignment="Left"
                            ToolTip="Add the selected parameters to the export">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE72A;" FontFamily="Segoe MDL2 Assets" FontSize="12"
                                       VerticalAlignment="Center" Margin="0,0,8,0" />
                            <TextBlock Text="Add" VerticalAlignment="Center" />
                        </StackPanel>
                    </Button>
                    <Button Name="RemoveParameterButton" HorizontalContentAlignment="Left"
                            ToolTip="Remove the selected columns from the export">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE72B;" FontFamily="Segoe MDL2 Assets" FontSize="12"
                                       VerticalAlignment="Center" Margin="0,0,8,0" />
                            <TextBlock Text="Remove" VerticalAlignment="Center" />
                        </StackPanel>
                    </Button>
                    <Separator Margin="0,14" />
                    <Button Name="MoveParameterUpButton" Margin="0,0,0,8" HorizontalContentAlignment="Left"
                            ToolTip="Move the selected columns earlier">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE74A;" FontFamily="Segoe MDL2 Assets" FontSize="12"
                                       VerticalAlignment="Center" Margin="0,0,8,0" />
                            <TextBlock Text="Move up" VerticalAlignment="Center" />
                        </StackPanel>
                    </Button>
                    <Button Name="MoveParameterDownButton" HorizontalContentAlignment="Left"
                            ToolTip="Move the selected columns later">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE74B;" FontFamily="Segoe MDL2 Assets" FontSize="12"
                                       VerticalAlignment="Center" Margin="0,0,8,0" />
                            <TextBlock Text="Move down" VerticalAlignment="Center" />
                        </StackPanel>
                    </Button>
                </StackPanel>

                <!-- EXPORTED -->
                <Border Grid.Column="4" Style="{StaticResource TableLane}" Margin="0"
                        Background="{StaticResource LaneGreenBrush}" Padding="12,0,12,12">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="EXPORTED COLUMNS" Style="{StaticResource TableGroupLabel}"
                                   Foreground="{StaticResource LaneGreenTextBrush}" Margin="0,10,0,10" />
                        <TextBlock Grid.Row="1" Style="{StaticResource CaptionText}" Margin="0,0,0,10"
                                   MinHeight="32" VerticalAlignment="Center"
                                   Text="Written left to right in this order. Id is always the first column." />
                        <ListBox Grid.Row="2" Name="SelectedParameterList" SelectionMode="Extended" MinHeight="120"
                                 AutomationProperties.Name="Exported columns" />
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <!-- Footer -->
        <DockPanel Grid.Row="3" Style="{StaticResource WwpFooter}">
            <Image Name="LogoImage" DockPanel.Dock="Left" Style="{StaticResource BrandLogoFooter}"
                   VerticalAlignment="Center" />
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
                <Button Name="BatchExportButton" Content="Save set" MinWidth="110" Margin="0,0,8,0"
                        Style="{StaticResource PrimaryButton}"
                        ToolTip="Save this set to Project Information" />
                <Button Name="OkButton" Content="Export" MinWidth="96" Margin="0,0,8,0"
                        ToolTip="Export this set now" />
                <Button Name="CancelButton" Content="Cancel" MinWidth="88" />
            </StackPanel>
        </DockPanel>
    </Grid>
</Window>'''




def _elem_id_int(eid):
    try:
        return int(eid.Value)      # Revit 2024+
    except AttributeError:
        return int(eid.Value)  # Revit 2023-

def sanitize_sheet_name(name):
    safe = re.sub(r"[:\\/?*\[\]]", "_", (name or "").strip())
    return (safe or "Schedule")[:31]


def _pluralize(name):
    if not name:
        return name
    parts = name.rsplit(" ", 1)
    last = parts[-1]
    lower_last = last.lower()
    if lower_last.endswith("s"):
        return name
    if lower_last.endswith(("x", "z", "ch", "sh")):
        last = last + "es"
    else:
        last = last + "s"
    parts[-1] = last
    return " ".join(parts)


def default_sheet_name_for_category(category_name):
    return sanitize_sheet_name(_pluralize(category_name or "Category Export"))


_PROFILE_TOKEN = "%userprofile%"


def _portable_path(path):
    """Replace the current user's profile folder (C:/Users/<name>) with %userprofile% so saved
    paths, notably ACC Desktop Connector (DC) folders, resolve for every user. Same rule as
    Export2Ex - By Schedule."""
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


def _same_saved_path(path):
    # Saved paths may predate the portable form; compare everything in that form.
    return _portable_path(_expand_path(path or ""))


def normalize_excel_output_path(path, default_ext=".xlsx"):
    value = (path or "").strip()
    if not value:
        return ""
    root, ext = os.path.splitext(value)
    if not ext:
        return value + default_ext
    if ext.lower() in ALLOWED_EXCEL_EXTENSIONS:
        return value
    return ""


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


def get_active_doc():
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc:
            return uidoc.Document
    except Exception:
        pass
    return None


def read_saved_sets(doc):
    try:
        proj_info = doc.ProjectInformation
        if proj_info is None:
            return {}
        param = proj_info.LookupParameter(PARAM_SAVED_SETS)
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


def write_saved_sets(doc, sets_dict):
    global _last_write_error
    _last_write_error = ""
    try:
        proj_info = doc.ProjectInformation
        if proj_info is None:
            _last_write_error = "ProjectInformation is None"
            return False
        param = proj_info.LookupParameter(PARAM_SAVED_SETS)
        t = DB.Transaction(doc, "Save Export2Ex Settings")
        t.Start()
        try:
            if param is None:
                app = doc.Application
                cat_set = app.Create.NewCategorySet()
                pi_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_ProjectInformation)
                if pi_cat is None:
                    t.RollBack()
                    _last_write_error = "Could not find Project Information category in document settings"
                    return False
                cat_set.Insert(pi_cat)
                binding = app.Create.NewInstanceBinding(cat_set)
                from Autodesk.Revit.DB import GroupTypeId
                opts = DB.InternalDefinitionCreationOptions(PARAM_SAVED_SETS, DB.SpecTypeId.String.Text)
                opts.Visible = True
                doc.ParameterBindings.Insert(opts, binding, GroupTypeId.Data)
                doc.Regenerate()
                proj_info = doc.ProjectInformation
                param = proj_info.LookupParameter(PARAM_SAVED_SETS)
            if param is None:
                t.RollBack()
                _last_write_error = "Parameter '{}' not found after creation attempt - it may already exist as a read-only shared parameter".format(PARAM_SAVED_SETS)
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
            param.Set(json.dumps(payload, ensure_ascii=False, indent=2, cls=_JsonEncoder))
            t.Commit()
            return True
        except Exception as inner:
            try:
                t.RollBack()
            except Exception:
                pass
            _last_write_error = str(inner)
            log_exception("write_saved_sets transaction", inner)
            return False
    except Exception as exc:
        _last_write_error = str(exc)
        log_exception("write_saved_sets", exc)
        return False


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


def add_lib_path():
    lib_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "lib"))
    if lib_path not in sys.path:
        sys.path.append(lib_path)


def load_uiutils():
    add_lib_path()
    import WWP_uiUtils as ui
    return ui


def get_config_and_saver():
    return get_tool_settings("Export2ExBeta", doc=get_active_doc())


def config_get(config, name, default=None):
    try:
        value = getattr(config, name)
    except Exception:
        return default
    return default if value is None else value


def _coerce_int(value, default=None):
    if value is None:
        return default
    if isinstance(value, (list, tuple)):
        for item in value:
            coerced = _coerce_int(item, None)
            if coerced is not None:
                return coerced
        return default
    try:
        return int(value)
    except Exception:
        pass
    text = str(value).strip()
    if not text:
        return default
    match = re.search(r"-?\d+", text)
    if not match:
        return default
    try:
        return int(match.group(0))
    except Exception:
        return default


def _coerce_string_list(value):
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [str(item) for item in value if str(item).strip()]
    text = str(value).strip()
    return [text] if text else []


def _normalize_mode(value):
    return MODE_BY_CATEGORY if value == MODE_BY_CATEGORY else MODE_FROM_SCHEDULE


def _pick_save_file(title, filter_text, default_extension, initial_directory, file_name):
    clr.AddReference("PresentationFramework")
    from Microsoft.Win32 import SaveFileDialog

    dialog = SaveFileDialog()
    dialog.Title = title or "Save File"
    dialog.Filter = filter_text or "All files (*.*)|*.*"
    if default_extension:
        dialog.DefaultExt = default_extension
        dialog.AddExtension = True
    if initial_directory and os.path.isdir(initial_directory):
        dialog.InitialDirectory = initial_directory
    if file_name:
        dialog.FileName = file_name
    result = dialog.ShowDialog()
    if result:
        return dialog.FileName
    return None


def get_default_dir(doc):
    if doc.IsWorkshared:
        try:
            central = doc.GetWorksharingCentralModelPath()
            if central:
                return os.path.dirname(DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(central))
        except Exception:
            pass
    if doc.PathName:
        return os.path.dirname(doc.PathName)
    return os.path.expanduser("~")


def ensure_existing_dir(path, fallback=""):
    if path and os.path.isdir(path):
        return path
    if fallback and os.path.isdir(fallback):
        return fallback
    return ""


def collect_schedules(doc):
    schedules = []
    for view in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule):
        if view.IsTemplate or view.IsTitleblockRevisionSchedule:
            continue
        schedules.append(view)
    schedules.sort(key=lambda item: item.Name)
    return schedules


def _get_linked_docs(doc):
    result = []
    try:
        for link in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
            link_doc = link.GetLinkDocument()
            if link_doc is not None:
                result.append(link_doc)
    except Exception:
        pass
    return result


def collect_category_records(doc):
    records = {}

    def _scan_doc(scan_doc):
        collector = DB.FilteredElementCollector(scan_doc).WhereElementIsNotElementType()
        for element in collector:
            try:
                category = element.Category
            except Exception:
                category = None
            if category is None:
                continue
            name = (category.Name or "").strip()
            if not name:
                continue
            category_id = element_id_value(category.Id)
            if category_id == -1:
                continue
            record = records.get(category_id)
            if record is None:
                record = {
                    "id": category.Id,
                    "id_value": category_id,
                    "name": name,
                    "count": 0,
                }
                records[category_id] = record
            record["count"] += 1

    _scan_doc(doc)
    for link_doc in _get_linked_docs(doc):
        _scan_doc(link_doc)

    result = list(records.values())
    result.sort(key=lambda item: item["name"].lower())
    return result


def get_elements_by_category(doc, category_id):
    def _collect(src_doc):
        try:
            collector = (
                DB.FilteredElementCollector(src_doc)
                .WherePasses(DB.ElementCategoryFilter(category_id))
                .WhereElementIsNotElementType()
            )
            return [e for e in collector.ToElements() if e is not None]
        except Exception:
            return []

    elements = _collect(doc)
    for link_doc in _get_linked_docs(doc):
        elements.extend(_collect(link_doc))
    return elements


def get_schedule_category_id(schedule):
    try:
        category_id = schedule.Definition.CategoryId
    except Exception:
        category_id = None
    if category_id is None or element_id_value(category_id) == -1:
        return None
    return category_id


def iter_parameter_names(element):
    if element is None:
        return
    for param in getattr(element, "Parameters", []):
        try:
            definition = param.Definition
            name = definition.Name if definition else None
        except Exception:
            name = None
        if name:
            yield name


def _iter_element_parameters(element):
    if element is None:
        return
    for param in getattr(element, "Parameters", []):
        try:
            definition = param.Definition
            name = definition.Name if definition else None
        except Exception:
            name = None
        if name:
            yield name, param


def get_parameter_options_for_category(doc, category_id, sample_limit=400):
    options = {}
    seen_types = set()
    for index, element in enumerate(get_elements_by_category(doc, category_id)):
        if index >= sample_limit:
            break
        for name, param in _iter_element_parameters(element):
            record = options.setdefault(name, {"name": name, "editable": False})
            try:
                if param is not None and not param.IsReadOnly:
                    record["editable"] = True
            except Exception:
                pass
        elem_type = get_element_type(doc, element)
        if elem_type is None:
            continue
        type_id = element_id_value(elem_type.Id)
        if type_id in seen_types:
            continue
        seen_types.add(type_id)
        for name, param in _iter_element_parameters(elem_type):
            record = options.setdefault(name, {"name": name, "editable": False})
            try:
                if param is not None and not param.IsReadOnly:
                    record["editable"] = True
            except Exception:
                pass
    try:
        proj_info = doc.ProjectInformation
        if proj_info is not None:
            for name, param in _iter_element_parameters(proj_info):
                record = options.setdefault(name, {"name": name, "editable": False})
                try:
                    if param is not None and not param.IsReadOnly:
                        record["editable"] = True
                except Exception:
                    pass
    except Exception:
        pass
    result = list(options.values())
    result.sort(key=lambda item: item["name"].lower())
    return result


class ScheduleItem(object):
    def __init__(self, view):
        self.view = view
        self.id_value = element_id_value(view.Id)
        self.category_id = get_schedule_category_id(view)
        self.display_name = "{} [id:{}]".format(view.Name, self.id_value).replace("_", "__")

    def __str__(self):
        return self.display_name


class CategoryItem(object):
    def __init__(self, record):
        self.record = record
        self.id_value = record["id_value"]
        self.display_name = "{} ({})".format(record["name"], record["count"]).replace("_", "__")

    def __str__(self):
        return self.display_name


def _to_net_list(values):
    result = List[String]()
    for value in values:
        result.Add("" if value is None else str(value))
    return result


def _to_net_object_list(values):
    result = List[Object]()
    for value in values:
        result.Add(value)
    return result


def _parameter_label(option):
    if option.get("editable", False):
        return option.get("name", "")
    return "{} (read-only)".format(option.get("name", ""))


def _make_parameter_list_item(option):
    item = ListBoxItem()
    item.Content = _parameter_label(option)
    item.Tag = option
    if not option.get("editable", False):
        item.ToolTip = "This parameter is read-only in the sampled category elements/types."
    return item


def _get_parameter_name_from_item(item):
    if item is None:
        return ""
    try:
        tag = item.Tag
    except Exception:
        tag = None
    if isinstance(tag, dict):
        return tag.get("name", "")
    return str(item)


def _load_export_window():
    from System.IO import StringReader
    from System.Windows.Markup import XamlReader
    from System.Xml import XmlReader

    xaml_path = os.path.join(os.path.dirname(__file__), "ExportSchedulesDialog.xaml")
    load_uiutils().uiUtils_ensure_theme()
    sources = [
        ("file", lambda: File.ReadAllText(xaml_path)),
        ("embedded", lambda: EMBEDDED_EXPORT_DIALOG_XAML),
    ]
    last_exc = None
    for source_name, source_loader in sources:
        try:
            xaml_text = source_loader()
            reader = XmlReader.Create(StringReader(xaml_text))
            window = XamlReader.Load(reader)
            if source_name != "file":
                log_message("Loaded Export2Ex Beta dialog from {} fallback".format(source_name))
            return window
        except Exception as exc:
            last_exc = exc
            log_exception("Failed to load Export2Ex Beta dialog from {}".format(source_name), exc)
    raise last_exc


def _export_sets_to_file(saved_sets, ui, tool_name="Export2ExBeta"):
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


SCHEDULE_TOOL_NAMESPACE = "export2ex"


def read_schedule_tool_sets(doc):
    """Saved sets of Export2Ex - By Schedule. They sit in the same Project Information
    parameter under their own namespace, or in the parameter that tool was pointed at."""
    param_names = [PARAM_SAVED_SETS]
    try:
        schedule_config, _save = get_tool_settings("Export2Ex", doc=doc)
        other = (getattr(schedule_config, "saved_sets_param_name", "") or "").strip()
        if other and other not in param_names:
            param_names.append(other)
    except Exception:
        pass
    for param_name in param_names:
        try:
            param = doc.ProjectInformation.LookupParameter(param_name)
            raw = (param.AsString() or "").strip() if param is not None else ""
            data = json.loads(raw) if raw else {}
        except Exception:
            continue
        block = data.get(SCHEDULE_TOOL_NAMESPACE) if isinstance(data, dict) else None
        if not isinstance(block, dict):
            continue
        # Current layout {"sets": ..., "settings": ...}; older projects keep the sets flat.
        sets = block.get("sets") if ("sets" in block or "settings" in block) else block
        if isinstance(sets, dict) and sets:
            return dict((name, data_) for name, data_ in sets.items() if isinstance(data_, dict))
    return {}


def _schedule_field_names(schedule):
    names = []
    for field_info in get_visible_schedule_fields(schedule):
        try:
            name = field_info["field"].GetName()
        except Exception:
            name = ""
        if name and name not in names:
            names.append(name)
    return names


def schedule_sets_to_category_sets(doc, schedule_sets, taken_names):
    """Turn By Schedule sets into "from schedule" sets: each schedule's category and visible
    fields, the set's Excel file, and the sheet name By Schedule would use. A set naming
    several schedules gives one set per schedule. Returns (new_sets, skipped_notes)."""
    schedules_by_name = dict((view.Name, view) for view in collect_schedules(doc))
    new_sets = {}
    notes = []
    for set_name in sorted(schedule_sets):
        sdata = schedule_sets[set_name]
        if _coerce_int(sdata.get("export_mode"), 0) != 0:
            notes.append("{}: a CSV set (By Category writes Excel only)".format(set_name))
            continue
        schedule_names = [n for n in (sdata.get("schedule_names") or []) if n]
        if not schedule_names:
            notes.append("{}: no schedule chosen".format(set_name))
            continue
        for schedule_name in schedule_names:
            new_name = set_name if len(schedule_names) == 1 else "{} - {}".format(set_name, schedule_name)
            view = schedules_by_name.get(schedule_name)
            category_id = get_schedule_category_id(view) if view is not None else None
            fields = _schedule_field_names(view) if category_id is not None else []
            if view is None:
                notes.append("{}: schedule '{}' is not in this model".format(new_name, schedule_name))
            elif category_id is None:
                notes.append("{}: '{}' is a multi-category schedule".format(new_name, schedule_name))
            elif not fields:
                notes.append("{}: '{}' has no visible fields".format(new_name, schedule_name))
            elif new_name in taken_names or new_name in new_sets:
                notes.append("{}: a By Category set with this name already exists".format(new_name))
            else:
                category = DB.Category.GetCategory(doc, category_id)
                if sdata.get("use_category_sheet_name", True) and category is not None:
                    sheet_name = default_sheet_name_for_category(category.Name)
                else:
                    sheet_name = sanitize_sheet_name(schedule_name)
                new_sets[new_name] = {
                    "mode": MODE_FROM_SCHEDULE,
                    "source_id": element_id_value(view.Id),
                    "category_id": element_id_value(category_id),
                    "param_names": fields,
                    "sheet_name": sheet_name,
                    "excel_path": _same_saved_path(normalize_excel_output_path(sdata.get("excel_path") or "")),
                }
    return new_sets, notes


def _load_batch_window():
    from System.IO import StringReader
    from System.Windows.Markup import XamlReader
    from System.Xml import XmlReader

    xaml_path = os.path.join(os.path.dirname(__file__), "BatchExportDialog.xaml")
    load_uiutils().uiUtils_ensure_theme()
    reader = XmlReader.Create(StringReader(File.ReadAllText(xaml_path)))
    return XamlReader.Load(reader)


def _show_beta_batch_dialog(saved_sets, doc, ui=None, add_callback=None, edit_callback=None):
    """Primary UI for Export2Ex Beta (BatchExportDialog.xaml): manage and batch-export saved sets.
    There is no Save button: sheet and path edits are written to Project Information once,
    when the window closes, and only if something changed. Add set and Edit go through the
    set editor, which saves the set itself."""
    clr.AddReference("PresentationFramework")
    clr.AddReference("PresentationCore")
    clr.AddReference("WindowsBase")
    from System.Windows import LogicalTreeHelper, Visibility
    from System.Windows.Controls.Primitives import PlacementMode

    paths = {}
    sheet_edits = {}
    _ok_clicked = [False]
    checkboxes = {}
    path_labels = {}
    sheet_boxes = {}

    window = _load_batch_window()
    apply_window_title(window, "Export2Ex Beta")
    window.Height = min(380 + max(len(saved_sets), 2) * 50, 720)

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
        shown = _portable_path(path)
        label.Text = shown or "Not set"
        label.ToolTip = _expand_path(path or "") or None
        label.Foreground = text_brush if shown else muted_brush

    def _category_name(sdata):
        cat_id = _coerce_int(sdata.get("category_id"), None)
        if cat_id in (None, -1):
            return ""
        try:
            category = DB.Category.GetCategory(doc, DB.ElementId(Int64(cat_id)))
            return category.Name if category is not None else ""
        except Exception:
            return ""

    def _source_summary(sdata):
        """Category name and field count shown in the SOURCE lane."""
        fields = len(sdata.get("param_names") or [])
        if sdata.get("mode") == MODE_FROM_SCHEDULE:
            origin = "from schedule"
        else:
            origin = "by category"
        field_text = "1 field" if fields == 1 else "{} fields".format(fields)
        return _category_name(sdata) or "No category", "{}, {}".format(field_text, origin)

    def _saved_sheet(sname, sdata):
        # A set without a sheet name exports to its category's plural, as the editor does.
        saved = (sdata.get("sheet_name") or "").strip()
        if saved:
            return saved
        category_name = _category_name(sdata)
        return default_sheet_name_for_category(category_name) if category_name else sanitize_sheet_name(sname)

    def _collect_sheet_edits():
        # Rows are rebuilt after Add, Edit and Delete; keep what was typed in the sheet boxes.
        for sname, box in list(sheet_boxes.items()):
            sheet_edits[sname] = (box.Text or "").strip()

    def _apply_edits():
        """Fold pending sheet and path edits into saved_sets (written by _write_sets)."""
        _collect_sheet_edits()
        for sname, sdata in list(saved_sets.items()):
            if not isinstance(sdata, dict):
                continue
            new_sheet = sheet_edits.get(sname, "")
            if new_sheet and new_sheet != _saved_sheet(sname, sdata):
                sdata["sheet_name"] = new_sheet
            new_path = (paths.get(sname) or "").strip()
            if new_path and new_path != _same_saved_path(sdata.get("excel_path")):
                sdata["excel_path"] = new_path

    def _snapshot():
        return json.dumps(saved_sets, sort_keys=True)

    # What Project Information holds; the window writes on close when saved_sets differs.
    _saved_snapshot = [_snapshot()]

    def _write_sets():
        ok_saved = write_saved_sets(doc, saved_sets)
        if ok_saved:
            _saved_snapshot[0] = _snapshot()
        return ok_saved

    def _forget_pending(names):
        # Sets replaced from outside this window drop their pending path and sheet edits.
        for name in names:
            paths.pop(name, None)
            sheet_edits.pop(name, None)

    def _make_edit_row(sname_):
        def _on_edit_row(_s, _e):
            sdata = dict(saved_sets.get(sname_) or {})
            # Hand the editor any output file chosen here but not saved yet.
            if paths.get(sname_):
                sdata["excel_path"] = paths[sname_]
            if edit_callback:
                edit_callback(sname_, sdata)
            _refresh_from_project()
        return _on_edit_row

    def _make_browse(sname_):
        def _on_browse(_s, _e):
            cur = _expand_path(paths.get(sname_) or "")
            init_dir = os.path.dirname(cur) if cur and os.path.isdir(os.path.dirname(cur)) else get_default_dir(doc)
            safe = re.sub(r'[\\/:*?"<>|]', "_", sname_)
            fname = os.path.basename(cur) if cur else "{}.xlsx".format(safe)
            new_path = _pick_save_file(
                title="'{}' -- Choose Output File".format(sname_),
                filter_text="Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                default_extension="xlsx",
                initial_directory=init_dir,
                file_name=fname,
            )
            new_path = normalize_excel_output_path(new_path or "")
            if new_path:
                paths[sname_] = _portable_path(new_path)
                _set_path_label(path_labels[sname_], paths[sname_])
        return _on_browse

    def _make_delete_row(sname_):
        # The set editor reads and writes the saved sets directly, so a delete is saved
        # at once rather than when the window closes.
        def _on_delete_row(_s, _e):
            _collect_sheet_edits()
            saved_sets.pop(sname_, None)
            paths.pop(sname_, None)
            sheet_edits.pop(sname_, None)
            _write_sets()
            _rebuild_rows()
        return _on_delete_row

    def _rebuild_rows():
        # Callers run _collect_sheet_edits() first, then drop what should not carry over.
        rows_host.Children.Clear()
        checkboxes.clear()
        path_labels.clear()
        sheet_boxes.clear()

        current_names = sorted(saved_sets.keys())
        set_count_text.Text = "1 set" if len(current_names) == 1 else "{} sets".format(len(current_names))
        empty_state.Visibility = Visibility.Collapsed if current_names else Visibility.Visible

        for sname in current_names:
            sdata = saved_sets.get(sname) or {}
            if sname not in paths:
                paths[sname] = _same_saved_path(sdata.get("excel_path"))

            row = row_template.LoadContent()

            checkboxes[sname] = _row_part(row, "RowCheck")

            name_tb = _row_part(row, "NameText")
            name_tb.Text = sname
            name_tb.ToolTip = sname

            source_name, fields_text = _source_summary(sdata)
            _row_part(row, "SourceText").Text = source_name
            _row_part(row, "FieldsText").Text = fields_text

            sheet_tb = _row_part(row, "SheetBox")
            sheet_tb.Text = sheet_edits.get(sname) or _saved_sheet(sname, sdata)
            sheet_boxes[sname] = sheet_tb

            path_tb = _row_part(row, "PathText")
            _set_path_label(path_tb, paths[sname])
            path_labels[sname] = path_tb

            _row_part(row, "BrowseButton").Click += _make_browse(sname)
            _row_part(row, "EditButton").Click += _make_edit_row(sname)
            _row_part(row, "DeleteButton").Click += _make_delete_row(sname)

            rows_host.Children.Add(row)

    _rebuild_rows()

    def _refresh_from_project():
        # The editor may have saved a set (Save set) or its last export path. What it saved
        # replaces this window's pending edit for that set; other pending edits are kept.
        _collect_sheet_edits()
        new_sets = read_saved_sets(doc)
        for name, data in new_sets.items():
            if not isinstance(data, dict):
                continue
            previous = saved_sets.get(name) or {}
            new_path = _same_saved_path(data.get("excel_path"))
            if new_path != _same_saved_path(previous.get("excel_path")):
                paths[name] = new_path
            if (data.get("sheet_name") or "").strip() != (previous.get("sheet_name") or "").strip():
                sheet_edits.pop(name, None)
        saved_sets.clear()
        saved_sets.update(new_sets)
        _saved_snapshot[0] = _snapshot()
        _rebuild_rows()

    def _add(_s, _e):
        if add_callback:
            add_callback()
        _refresh_from_project()

    def _ok(_s, _e):
        _ok_clicked[0] = True
        window.Close()

    sets_file_button = window.FindName("SetsFileButton")

    def _open_sets_menu(_s, _e):
        menu = sets_file_button.ContextMenu
        menu.PlacementTarget = sets_file_button
        menu.Placement = PlacementMode.Bottom
        menu.IsOpen = True

    def _save_file(_s, _e):
        _apply_edits()
        _export_sets_to_file(saved_sets, ui, tool_name="Export2ExBeta")

    def _load_file(_s, _e):
        imported = _import_sets_from_file(ui)
        if imported is not None:
            _apply_edits()
            _forget_pending(imported.keys())
            saved_sets.update(imported)
            _write_sets()
            _rebuild_rows()

    def _copy_schedule_sets(_s, _e):
        # Reuse the By Schedule sets as a starting point; Edit then adds parameters.
        dialogs = ui or load_uiutils()
        schedule_sets = read_schedule_tool_sets(doc)
        if not schedule_sets:
            dialogs.uiUtils_alert("No sets saved in Export2Ex - By Schedule were found in this project.",
                                  title="Export2Ex Beta")
            return
        names = sorted(schedule_sets)
        labels = [
            "{}   ({})".format(name, ", ".join(schedule_sets[name].get("schedule_names") or []) or "no schedule")
            for name in names
        ]
        picked = dialogs.uiUtils_select_indices(
            labels,
            title="Copy sets from Export by Schedule",
            prompt="Pick the By Schedule sets to copy. Each becomes a set with its schedule's category, "
                   "visible fields and Excel file; choose Edit afterwards to add more parameters.",
            multiselect=True,
            width=760,
            height=520,
        )
        if not picked:
            return
        chosen = dict((names[i], schedule_sets[names[i]]) for i in picked)
        _apply_edits()
        new_sets, notes = schedule_sets_to_category_sets(doc, chosen, set(saved_sets.keys()))
        if new_sets:
            _forget_pending(new_sets.keys())
            saved_sets.update(new_sets)
            if not _write_sets():
                notes.append("Could not save to Project Information: {}".format(_last_write_error))
            _rebuild_rows()
        message = "Copied {} set(s).".format(len(new_sets))
        if notes:
            message += "\n\nNot copied:\n" + "\n".join(notes)
        dialogs.uiUtils_alert(message, title="Export2Ex Beta")

    # CancelButton ("Close") is IsCancel in the XAML: it closes the window (and so does Esc).
    window.FindName("AddSetButton").Click += _add
    window.FindName("CopyScheduleSetsButton").Click += _copy_schedule_sets
    window.FindName("ExportSelectedButton").Click += _ok
    sets_file_button.Click += _open_sets_menu
    window.FindName("SaveFileMenuItem").Click += _save_file
    window.FindName("LoadFileMenuItem").Click += _load_file

    window.ShowDialog()

    _apply_edits()
    if _snapshot() != _saved_snapshot[0] and not _write_sets() and ui:
        ui.uiUtils_alert(
            "Could not save the sheet and path changes to Project Information.\n{}".format(_last_write_error),
            title="Export2Ex Beta",
        )

    if not _ok_clicked[0]:
        return None

    names = sorted(saved_sets.keys())
    selected = [sname for sname in names if sname in checkboxes and checkboxes[sname].IsChecked]
    final_sheet_names = dict(
        (sname, sheet_edits.get(sname) or _saved_sheet(sname, saved_sets[sname]))
        for sname in names if isinstance(saved_sets[sname], dict)
    )
    return selected, dict(paths), final_sheet_names


def show_export_form(ui, doc, schedules, categories, init_excel_path, initial_mode, initial_source_id, initial_category_id, initial_param_names, initial_sheet_name="", initial_set_name=""):
    clr.AddReference("PresentationFramework")
    clr.AddReference("PresentationCore")
    clr.AddReference("WindowsBase")

    window = _load_export_window()
    apply_window_title(window, "Export2Ex Beta")

    from_schedule_toggle = window.FindName("FromScheduleToggle")
    units_box = window.FindName("UnitsBox")
    source_search_box = window.FindName("SourceSearchBox")
    source_label = window.FindName("SourceLabel")
    source_list_label = window.FindName("SourceListLabel")
    source_list = window.FindName("SourceList")
    parameter_search_box = window.FindName("ParameterSearchBox")
    parameter_list = window.FindName("ParameterList")
    selected_parameter_list = window.FindName("SelectedParameterList")
    add_parameter_button = window.FindName("AddParameterButton")
    remove_parameter_button = window.FindName("RemoveParameterButton")
    move_parameter_up_button = window.FindName("MoveParameterUpButton")
    move_parameter_down_button = window.FindName("MoveParameterDownButton")
    ok_button = window.FindName("OkButton")
    cancel_button = window.FindName("CancelButton")
    batch_export_button = window.FindName("BatchExportButton")
    logo_image = window.FindName("LogoImage")
    show_read_only_filter = window.FindName("ShowReadOnlyFilter")
    sheet_name_box = window.FindName("SheetNameBox")
    set_name_box = window.FindName("SetNameBox")

    schedule_items = schedules or []
    category_items = categories or []
    parameter_cache = {}
    selected_params_by_category = {}
    initial_category_id = _coerce_int(initial_category_id, None)
    initial_source_id = _coerce_int(initial_source_id, None)
    initial_param_names = _coerce_string_list(initial_param_names)
    if initial_category_id is not None and initial_param_names:
        selected_params_by_category[initial_category_id] = list(initial_param_names)

    def _current_mode():
        try:
            return MODE_FROM_SCHEDULE if from_schedule_toggle.IsChecked else MODE_BY_CATEGORY
        except Exception:
            return MODE_BY_CATEGORY

    from_schedule_toggle.IsChecked = (initial_mode == MODE_FROM_SCHEDULE)

    def _get_source_items():
        return category_items if _current_mode() == MODE_BY_CATEGORY else schedule_items

    def _resolve_category_id(item):
        if item is None:
            return None
        if _current_mode() == MODE_BY_CATEGORY:
            return item.id_value
        if item.category_id is None:
            return None
        return element_id_value(item.category_id)

    def _category_name_for_id(category_id):
        if category_id is None:
            return ""
        for item in category_items:
            if item.id_value == category_id:
                return item.record["name"]
        return ""

    def _default_sheet_name_for_item(item):
        category_id = _resolve_category_id(item)
        category_name = _category_name_for_id(category_id)
        if category_name:
            return default_sheet_name_for_category(category_name)
        if item is not None and _current_mode() == MODE_BY_CATEGORY:
            return default_sheet_name_for_category(item.record["name"])
        return default_sheet_name_for_category("Category Export")

    def _get_parameter_names(category_id):
        if category_id is None:
            return []
        if category_id not in parameter_cache:
            parameter_cache[category_id] = get_parameter_options_for_category(doc, DB.ElementId(Int64(category_id)))
        return parameter_cache[category_id]

    def _get_selected_parameter_names(category_id):
        if category_id is None:
            return []
        selected = selected_params_by_category.get(category_id)
        if selected is None:
            selected = []
            selected_params_by_category[category_id] = selected
        return selected

    def _refresh_selected_parameter_list(category_id, preserve_selection=None):
        selected_names = _get_selected_parameter_names(category_id)
        selected_parameter_list.Items.Clear()
        id_item = ListBoxItem()
        id_item.Content = "Id  -- always exported"
        id_item.Tag = {"name": "__id__", "editable": False, "__permanent__": True}
        id_item.Foreground = window.TryFindResource("AccentBrush") or Brushes.SteelBlue
        selected_parameter_list.Items.Add(id_item)
        option_map = dict((option["name"], option) for option in _get_parameter_names(category_id))
        for name in selected_names:
            option = option_map.get(name, {"name": name, "editable": True})
            selected_parameter_list.Items.Add(_make_parameter_list_item(option))
        preserve = list(preserve_selection or [])
        try:
            selected_parameter_list.SelectedItems.Clear()
            for item in selected_parameter_list.Items:
                if _get_parameter_name_from_item(item) in preserve:
                    selected_parameter_list.SelectedItems.Add(item)
        except Exception:
            pass

    def _refresh_parameter_list():
        selected_item = source_list.SelectedItem
        category_id = _resolve_category_id(selected_item)
        all_options = _get_parameter_names(category_id)
        selected_names = set(_get_selected_parameter_names(category_id))
        text = (parameter_search_box.Text or "").strip().lower()
        filtered = [option for option in all_options if option["name"] not in selected_names]
        try:
            if not show_read_only_filter.IsChecked:
                filtered = [option for option in filtered if option.get("editable", False)]
        except Exception:
            pass
        if text:
            filtered = [option for option in filtered if text in option["name"].lower()]
        parameter_list.Items.Clear()
        for option in filtered:
            parameter_list.Items.Add(_make_parameter_list_item(option))
        _refresh_selected_parameter_list(category_id)

    def _refresh_source_list():
        mode = _current_mode()
        items = _get_source_items()
        text = (source_search_box.Text or "").strip().lower()
        filtered = items if not text else [item for item in items if text in item.display_name.lower()]
        source_list.DisplayMemberPath = "display_name"
        source_list.ItemsSource = _to_net_object_list(filtered)
        source_label.Text = "Search categories" if mode == MODE_BY_CATEGORY else "Search schedules"
        source_list_label.Text = "Categories" if mode == MODE_BY_CATEGORY else "Schedules"
        target_id = initial_source_id
        selected = None
        for item in filtered:
            item_id = item.id_value if mode == MODE_BY_CATEGORY else item.id_value
            if target_id is not None and item_id == target_id:
                selected = item
                break
        if selected is None and filtered:
            selected = filtered[0]
        source_list.SelectedItem = selected
        _refresh_parameter_list()

    _initialized = [False]
    _sheet_name_autofill = [True]
    _sheet_name_updating = [False]
    _last_auto_sheet_name = [""]

    def _set_sheet_name_text(value):
        if sheet_name_box is None:
            return
        _sheet_name_updating[0] = True
        try:
            sheet_name_box.Text = value or ""
        finally:
            _sheet_name_updating[0] = False

    def _refresh_sheet_name_default(force=False):
        if sheet_name_box is None:
            return
        selected_item = source_list.SelectedItem
        auto_name = _default_sheet_name_for_item(selected_item)
        current = (sheet_name_box.Text or "").strip()
        if force or _sheet_name_autofill[0] or not current or current == _last_auto_sheet_name[0]:
            _last_auto_sheet_name[0] = auto_name
            _sheet_name_autofill[0] = True
            _set_sheet_name_text(auto_name)

    def _sheet_name_changed(_sender=None, _args=None):
        if _sheet_name_updating[0] or sheet_name_box is None:
            return
        current = (sheet_name_box.Text or "").strip()
        _sheet_name_autofill[0] = not current or current == _last_auto_sheet_name[0]

    def _auto_populate_schedule_fields():
        if not _initialized[0] or _current_mode() != MODE_FROM_SCHEDULE:
            return
        selected_item = source_list.SelectedItem
        if selected_item is None:
            return
        category_id = _resolve_category_id(selected_item)
        if category_id is None:
            return
        try:
            fields = get_visible_schedule_fields(selected_item.view)
        except Exception:
            return
        names = []
        for f in fields:
            try:
                name = f["field"].GetName()
                if name and name not in names:
                    names.append(name)
            except Exception:
                pass
        selected_params_by_category[category_id] = names

    def _source_selection_changed(_sender, _args):
        _auto_populate_schedule_fields()
        _refresh_parameter_list()
        _refresh_sheet_name_default()

    def _add_parameters(_sender=None, _args=None):
        selected_item = source_list.SelectedItem
        category_id = _resolve_category_id(selected_item)
        if category_id is None:
            return
        selected_names = _get_selected_parameter_names(category_id)
        for item in parameter_list.SelectedItems:
            name = _get_parameter_name_from_item(item)
            if name not in selected_names:
                selected_names.append(name)
        _refresh_parameter_list()

    def _remove_parameters(_sender=None, _args=None):
        selected_item = source_list.SelectedItem
        category_id = _resolve_category_id(selected_item)
        if category_id is None:
            return
        selected_names = _get_selected_parameter_names(category_id)
        removed = set(
            _get_parameter_name_from_item(item)
            for item in selected_parameter_list.SelectedItems
            if not (isinstance(getattr(item, "Tag", None), dict) and item.Tag.get("__permanent__"))
        )
        if not removed:
            return
        selected_params_by_category[category_id] = [name for name in selected_names if name not in removed]
        _refresh_parameter_list()

    def _move_selected_parameters(direction):
        selected_item = source_list.SelectedItem
        category_id = _resolve_category_id(selected_item)
        if category_id is None:
            return
        selected_names = _get_selected_parameter_names(category_id)
        selected_now = [_get_parameter_name_from_item(item) for item in selected_parameter_list.SelectedItems]
        if not selected_now:
            return
        indices = [idx for idx, name in enumerate(selected_names) if name in selected_now]
        if direction < 0:
            for idx in indices:
                if idx <= 0:
                    continue
                selected_names[idx - 1], selected_names[idx] = selected_names[idx], selected_names[idx - 1]
        else:
            for idx in reversed(indices):
                if idx >= len(selected_names) - 1:
                    continue
                selected_names[idx + 1], selected_names[idx] = selected_names[idx], selected_names[idx + 1]
        _refresh_selected_parameter_list(category_id, preserve_selection=selected_now)

    def _move_up(_sender=None, _args=None):
        _move_selected_parameters(-1)

    def _move_down(_sender=None, _args=None):
        _move_selected_parameters(1)

    saved_sets = read_saved_sets(doc)

    def _get_current_set_data():
        selected_item = source_list.SelectedItem
        category_id = _resolve_category_id(selected_item)
        source_id = selected_item.id_value if selected_item is not None else None
        return {
            "mode": _current_mode(),
            "source_id": int(source_id) if source_id is not None else None,
            "category_id": int(category_id) if category_id is not None else None,
            "param_names": list(selected_params_by_category.get(category_id, [])),
            "sheet_name": (sheet_name_box.Text or "").strip() if sheet_name_box is not None else "",
        }

    def _apply_saved_set(set_data):
        _initialized[0] = False
        try:
            mode = set_data.get("mode", MODE_BY_CATEGORY)
            source_id = _coerce_int(set_data.get("source_id"), None)
            category_id = _coerce_int(set_data.get("category_id"), None)
            param_names = set_data.get("param_names") or []
            saved_sheet = (set_data.get("sheet_name") or "").strip()
            saved_excel = (set_data.get("excel_path") or "").strip()
            if category_id is not None:
                selected_params_by_category[category_id] = list(param_names)
            from_schedule_toggle.IsChecked = (mode == MODE_FROM_SCHEDULE)
            source_search_box.Text = ""
            for item in _get_source_items():
                if item.id_value == source_id:
                    source_list.SelectedItem = item
                    break
            if sheet_name_box is not None and saved_sheet:
                sheet_name_box.Text = saved_sheet
            _refresh_parameter_list()
        finally:
            _initialized[0] = True

    _batch_result = [None]

    def _save_set(_sender=None, _args=None):
        name = (set_name_box.Text or "").strip() if set_name_box is not None else ""
        if not name:
            ui.uiUtils_alert("Enter a name in the Set name field before saving.", title="Export2Ex Beta")
            return
        new_data = _get_current_set_data()
        existing = saved_sets.get(name) or {}
        if existing.get("excel_path"):
            new_data["excel_path"] = existing["excel_path"]
        saved_sets[name] = new_data
        if not write_saved_sets(doc, saved_sets):
            ui.uiUtils_alert(
                "Could not write to '{}' on Project Information.\nCheck the parameter exists and is not read-only.".format(PARAM_SAVED_SETS),
                title="Export2Ex Beta",
            )
            return
        ui.uiUtils_alert("Set '{}' saved.".format(name), title="Export2Ex Beta")

    _chosen_excel_path = [""]

    def _ok(_sender, _args):
        # The editor has no path box, so Export asks where to write. Check the inputs
        # first and treat a cancelled picker as "not yet": the editor stays open.
        category_id = _resolve_category_id(source_list.SelectedItem)
        if category_id in (None, -1):
            ui.uiUtils_alert("Select a schedule or category with a valid category.", title="Export2Ex Beta")
            return
        if not selected_params_by_category.get(category_id):
            ui.uiUtils_alert("Select at least one parameter to export.", title="Export2Ex Beta")
            return
        # Start from the set's own file (the batch window passes in an unsaved Browse
        # choice); a Set name typed here that matches another saved set starts from that.
        name = (set_name_box.Text or "").strip() if set_name_box is not None else ""
        seed = init_excel_path
        if name and name != initial_set_name:
            seed = (saved_sets.get(name) or {}).get("excel_path") or seed
        seed = normalize_excel_output_path(_expand_path(seed or ""))
        seed_dir = os.path.dirname(seed) if seed else ""
        picked = _pick_save_file(
            title="Export2Ex Beta -- Choose Output File",
            filter_text="Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            default_extension=os.path.splitext(seed)[1].lstrip(".") if seed else "xlsx",
            initial_directory=seed_dir if os.path.isdir(seed_dir) else get_default_dir(doc),
            file_name=os.path.basename(seed) if seed else "CategoryExport.xlsx",
        )
        if not picked:
            return
        picked = normalize_excel_output_path(picked)
        if not picked:
            ui.uiUtils_alert("Choose an Excel file path ending with .xlsx or .xlsm.", title="Export2Ex Beta")
            return
        _chosen_excel_path[0] = picked
        window.DialogResult = True
        window.Close()

    def _cancel(_sender, _args):
        window.DialogResult = False
        window.Close()

    ui.uiUtils_load_logo(logo_image)

    from_schedule_toggle.Checked += lambda _sender, _args: _refresh_source_list()
    from_schedule_toggle.Unchecked += lambda _sender, _args: _refresh_source_list()
    source_search_box.TextChanged += lambda _sender, _args: _refresh_source_list()
    source_list.SelectionChanged += _source_selection_changed
    parameter_search_box.TextChanged += lambda _sender, _args: _refresh_parameter_list()
    show_read_only_filter.Checked += lambda _sender, _args: _refresh_parameter_list()
    show_read_only_filter.Unchecked += lambda _sender, _args: _refresh_parameter_list()
    add_parameter_button.Click += _add_parameters
    remove_parameter_button.Click += _remove_parameters
    move_parameter_up_button.Click += _move_up
    move_parameter_down_button.Click += _move_down
    ok_button.Click += _ok
    cancel_button.Click += _cancel
    if batch_export_button is not None:
        batch_export_button.Click += _save_set
    if sheet_name_box is not None:
        sheet_name_box.TextChanged += _sheet_name_changed
    if set_name_box is not None and initial_set_name:
        set_name_box.Text = initial_set_name
    _refresh_source_list()
    _initialized[0] = True

    if sheet_name_box is not None:
        initial_sheet_name_stripped = (initial_sheet_name or "").strip()
        if initial_sheet_name_stripped:
            _last_auto_sheet_name[0] = _default_sheet_name_for_item(source_list.SelectedItem)
            _sheet_name_autofill[0] = initial_sheet_name_stripped == _last_auto_sheet_name[0]
            _set_sheet_name_text(initial_sheet_name_stripped)
        else:
            _refresh_sheet_name_default(force=True)

    if not window.ShowDialog():
        if _batch_result[0] is not None:
            selected, paths, sheet_names = _batch_result[0]
            return {"batch_sets": selected, "batch_paths": paths, "batch_sheet_names": sheet_names}
        return None

    selected_item = source_list.SelectedItem
    category_id = _resolve_category_id(selected_item)
    source_id = None
    source_name = ""
    if selected_item is not None:
        if _current_mode() == MODE_BY_CATEGORY:
            source_id = selected_item.id_value
            source_name = selected_item.record["name"]
        else:
            source_id = selected_item.id_value
            source_name = selected_item.view.Name
    units_mode = "project"
    try:
        if units_box is not None:
            idx = units_box.SelectedIndex
            if idx == 1:
                units_mode = "imperial"
            elif idx == 2:
                units_mode = "metric"
    except Exception:
        pass
    return {
        "mode": _current_mode(),
        "source_id": source_id,
        "source_name": source_name,
        "category_id": category_id,
        "selected_param_names": list(selected_params_by_category.get(category_id, [])),
        "excel_path": _chosen_excel_path[0],
        "set_name": (set_name_box.Text or "").strip() if set_name_box is not None else "",
        "sheet_name": (sheet_name_box.Text or "").strip() if sheet_name_box is not None else "",
        "units_mode": units_mode,
    }


def get_section_data(schedule, section_type):
    try:
        return schedule.GetTableData().GetSectionData(section_type)
    except Exception:
        return None


def get_cell_text(schedule, section_type, row, col):
    try:
        return schedule.GetCellText(section_type, row, col) or ""
    except Exception:
        pass
    try:
        section = get_section_data(schedule, section_type)
        if section is not None:
            return section.GetCellText(row, col) or ""
    except Exception:
        pass
    return ""


def get_body_row_element_ids(schedule):
    section = get_section_data(schedule, DB.SectionType.Body)
    if section is None:
        return []
    ids = []
    row_count = int(getattr(section, "NumberOfRows", 0))
    col_count = int(getattr(section, "NumberOfColumns", 0))
    for row in range(row_count):
        found = -1
        for col in range(col_count):
            try:
                elem_id = section.GetCellElementId(row, col)
            except Exception:
                elem_id = None
            found = element_id_value(elem_id)
            if found != -1:
                break
        ids.append(found)
    return ids


def collect_schedule_elements(schedule):
    try:
        elements = list(
            DB.FilteredElementCollector(schedule.Document, schedule.Id)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        elements = []
    return [elem for elem in elements if elem is not None]


def get_visible_schedule_fields(schedule):
    fields = []
    try:
        definition = schedule.Definition
        field_ids = list(definition.GetFieldOrder())
    except Exception:
        return fields
    for col_index, field_id in enumerate(field_ids):
        try:
            field = definition.GetField(field_id)
        except Exception:
            continue
        if field is None:
            continue
        try:
            if field.IsHidden:
                continue
        except Exception:
            pass
        try:
            heading = (field.ColumnHeading or "").strip()
        except Exception:
            heading = ""
        if not heading:
            try:
                heading = (field.GetName() or "").strip()
            except Exception:
                heading = ""
        fields.append(
            {
                "field": field,
                "column_index": len(fields),
                "heading": heading or "Column {}".format(col_index + 1),
                "param_id": safe_field_parameter_id(field),
            }
        )
    return fields


def safe_field_parameter_id(field):
    try:
        param_id = field.ParameterId
    except Exception:
        param_id = None
    if param_id is not None and element_id_value(param_id) == -1:
        return None
    return param_id


def is_key_schedule(view):
    try:
        definition = view.Definition
    except Exception:
        definition = None
    try:
        return bool(definition and definition.IsKeySchedule)
    except Exception:
        return False


def get_element_type(doc, element):
    try:
        type_id = element.GetTypeId()
    except Exception:
        type_id = None
    if type_id is None or element_id_value(type_id) == -1:
        return None
    try:
        return doc.GetElement(type_id)
    except Exception:
        return None


def get_parameter_from_element_or_type(doc, element, param_id):
    if element is None or param_id is None:
        return None
    try:
        param = element.get_Parameter(param_id)
    except Exception:
        param = None
    if param:
        return param
    elem_type = get_element_type(doc, element)
    if elem_type is None:
        return None
    try:
        return elem_type.get_Parameter(param_id)
    except Exception:
        return None


def parameter_to_text(doc, param):
    if not param:
        return ""
    try:
        value = param.AsValueString()
        if value not in (None, ""):
            return value
    except Exception:
        pass
    try:
        storage = param.StorageType
    except Exception:
        storage = None
    try:
        if storage == DB.StorageType.String:
            return param.AsString() or ""
        if storage == DB.StorageType.Integer:
            return str(param.AsInteger())
        if storage == DB.StorageType.Double:
            return str(param.AsDouble())
        if storage == DB.StorageType.ElementId:
            ref_id = param.AsElementId()
            ref_val = element_id_value(ref_id)
            if ref_val == -1:
                return ""
            ref_elem = doc.GetElement(ref_id)
            if ref_elem is not None:
                for attr in ("Name",):
                    try:
                        value = getattr(ref_elem, attr)
                        if value:
                            return str(value)
                    except Exception:
                        pass
            return str(ref_val)
    except Exception:
        pass
    return ""


def get_parameter_by_name(doc, element, param_name):
    if element is None or not param_name:
        return None
    try:
        param = element.LookupParameter(param_name)
    except Exception:
        param = None
    if param:
        return param
    elem_type = get_element_type(doc, element)
    if elem_type is not None:
        try:
            param = elem_type.LookupParameter(param_name)
        except Exception:
            param = None
        if param:
            return param
    try:
        proj_info = doc.ProjectInformation
        if proj_info is not None:
            return proj_info.LookupParameter(param_name)
    except Exception:
        pass
    return None


def _pick_unit_for_spec(spec_type_id, prefer_imperial):
    try:
        valid = list(DB.UnitUtils.GetValidUnits(spec_type_id))
        if not valid:
            return None
        def _uid_str(uid):
            try:
                return uid.TypeId.lower()
            except Exception:
                return str(uid).lower()
        if prefer_imperial:
            keywords = ["feet", "foot", "inch", "mile", "acre"]
            for uid in valid:
                if any(k in _uid_str(uid) for k in keywords):
                    return uid
        else:
            seeks = ["meters", "metre"]
            skip = ["milli", "centi", "kilo", "micro"]
            for uid in valid:
                s = _uid_str(uid)
                if any(k in s for k in seeks) and not any(k in s for k in skip):
                    return uid
            for uid in valid:
                if "milli" in _uid_str(uid):
                    return uid
        return None
    except Exception:
        return None


def parameter_to_export_value(doc, param, units_mode="project"):
    if not param:
        return ""
    try:
        storage = param.StorageType
    except Exception:
        storage = None
    try:
        if storage == DB.StorageType.String:
            return param.AsString() or ""
        if storage == DB.StorageType.Double:
            raw = param.AsDouble()
            if units_mode == "project":
                # param.GetUnitTypeId() returns the project's display unit for this param
                try:
                    unit_type_id = param.GetUnitTypeId()
                    tid = ""
                    try:
                        tid = unit_type_id.TypeId or ""
                    except Exception:
                        pass
                    if tid:
                        return DB.UnitUtils.ConvertFromInternalUnits(raw, unit_type_id)
                except Exception:
                    pass
            elif units_mode in ("imperial", "metric"):
                spec_type_id = None
                try:
                    spec_type_id = param.Definition.GetSpecTypeId()
                except Exception:
                    pass
                if spec_type_id is not None:
                    try:
                        if DB.UnitUtils.IsMeasurableSpec(spec_type_id):
                            display_unit = _pick_unit_for_spec(spec_type_id, prefer_imperial=(units_mode == "imperial"))
                            if display_unit is not None:
                                return DB.UnitUtils.ConvertFromInternalUnits(raw, display_unit)
                    except Exception:
                        pass
            return raw
        if storage == DB.StorageType.Integer:
            value_string = param.AsValueString()
            if value_string not in (None, "") and value_string not in ("0", "1"):
                return value_string
            return param.AsInteger()
        if storage == DB.StorageType.ElementId:
            ref_id = param.AsElementId()
            ref_value = element_id_value(ref_id)
            if ref_value == -1:
                return ""
            ref_elem = doc.GetElement(ref_id)
            if ref_elem is not None:
                try:
                    name = getattr(ref_elem, "Name", None)
                    if name:
                        return name
                except Exception:
                    pass
            return str(ref_value)
    except Exception:
        pass
    try:
        value = param.AsValueString()
        if value not in (None, ""):
            return value
    except Exception:
        pass
    return ""


def build_category_export_rows(doc, category_id, param_names, units_mode="project"):
    elements = get_elements_by_category(doc, category_id)
    elements.sort(key=lambda item: element_id_value(item.Id))
    headers = ["Id"] + list(param_names or [])
    rows = []
    for element in elements:
        row = [element_id_value(element.Id)]
        for param_name in param_names or []:
            row.append(parameter_to_export_value(doc, get_parameter_by_name(doc, element, param_name), units_mode=units_mode))
        rows.append(row)
    return headers, rows, len(elements)


def build_schedule_body_rows(schedule):
    section = get_section_data(schedule, DB.SectionType.Body)
    if section is None:
        return []
    row_count = int(getattr(section, "NumberOfRows", 0))
    col_count = int(getattr(section, "NumberOfColumns", 0))
    rows = []
    for row in range(row_count):
        rows.append([get_cell_text(schedule, DB.SectionType.Body, row, col) for col in range(col_count)])
    return rows


def build_schedule_row_lookup(schedule):
    row_ids = get_body_row_element_ids(schedule)
    row_values = build_schedule_body_rows(schedule)
    lookup = {}
    for row_id, values in zip(row_ids, row_values):
        if row_id == -1:
            continue
        lookup.setdefault(row_id, []).append(values)
    return row_ids, lookup


def build_table_headers(schedule, column_count):
    section = get_section_data(schedule, DB.SectionType.Header)
    row_count = int(getattr(section, "NumberOfRows", 0)) if section is not None else 0
    if row_count > 0:
        last_row = row_count - 1
        headers = [get_cell_text(schedule, DB.SectionType.Header, last_row, col) for col in range(column_count)]
        if any(header.strip() for header in headers):
            return [header or "Column {}".format(idx + 1) for idx, header in enumerate(headers)]
    return ["Column {}".format(idx + 1) for idx in range(column_count)]


def order_schedule_elements(schedule, elements, row_ids):
    elem_by_id = {}
    ordered = []
    seen = set()
    for elem in elements:
        elem_by_id[element_id_value(elem.Id)] = elem
    for row_id in row_ids:
        if row_id in seen:
            continue
        elem = elem_by_id.get(row_id)
        if elem is not None:
            ordered.append(elem)
            seen.add(row_id)
    for elem in sorted(elements, key=lambda item: element_id_value(item.Id)):
        elem_id = element_id_value(elem.Id)
        if elem_id not in seen:
            ordered.append(elem)
            seen.add(elem_id)
    return ordered


def get_field_text_for_element(doc, schedule, element, field_info, row_lookup):
    param_id = field_info.get("param_id")
    value = ""
    if param_id is not None:
        value = parameter_to_text(doc, get_parameter_from_element_or_type(doc, element, param_id))
    if value not in (None, ""):
        return value
    row_queue = row_lookup.get(element_id_value(element.Id)) or []
    if row_queue:
        column_index = int(field_info.get("column_index", 0))
        if column_index < len(row_queue[0]):
            return row_queue[0][column_index]
    return ""


def _inject_element_ids(headers, rows, ids):
    """Prepend an 'Element ID' column to headers and rows.

    `ids` is a parallel list of integer element IDs (or -1 for rows without
    a traceable element, e.g. group-header rows).  Rows with id == -1 get an
    empty cell in the Element ID column.
    """
    new_headers = ["Element ID"] + list(headers)
    new_rows = []
    for i, row in enumerate(rows):
        eid = ids[i] if i < len(ids) else -1
        new_rows.append(([str(eid)] if eid != -1 else [""]) + list(row))
    return new_headers, new_rows


def strip_leading_header_and_blanks(body_rows, headers=None):
    """Remove duplicate column-header rows and blank separator rows from the
    top of the body data.

    Some schedule types (key schedules, schedules with the 'blank row before
    data' appearance option) include the column-header row and/or a blank row
    as the first rows of the body section.  Because the exporter already writes
    its own bold header row, those body rows must be dropped to avoid
    duplication and empty rows in the output.
    """
    if not body_rows:
        return body_rows
    result = list(body_rows)
    headers_lower = [h.strip().lower() for h in headers] if headers else []
    while result:
        row = result[0]
        row_lower = [cell.strip().lower() for cell in row]
        if headers_lower and row_lower == headers_lower:
            result.pop(0)
        elif all(cell == "" for cell in row):
            result.pop(0)
        else:
            break
    return result


def _strip_body_rows(body_rows, headers, raw_ids):
    """Strip leading header/blank rows and keep raw_ids in sync.

    Returns (stripped_rows, aligned_ids) where aligned_ids[i] is the element
    ID that corresponds to stripped_rows[i].
    """
    n_before = len(body_rows)
    stripped = strip_leading_header_and_blanks(body_rows, headers)
    n_stripped = n_before - len(stripped)
    aligned_ids = (raw_ids[n_stripped:n_stripped + len(stripped)]
                   if raw_ids else [-1] * len(stripped))
    return stripped, aligned_ids


def build_schedule_export_rows(doc, schedule):
    fields = get_visible_schedule_fields(schedule)
    elements = collect_schedule_elements(schedule)
    key_sched = is_key_schedule(schedule)

    # GetCellText-based approach is the most reliable primary source because it
    # reflects what Revit actually renders in the schedule, covering calculated
    # value fields and schedules where GetCellElementId returns invalid IDs
    # (grouped schedules, key schedules, etc.).
    body_rows = build_schedule_body_rows(schedule)
    # Fetch element IDs now, before any stripping, so indices stay aligned.
    raw_ids = [] if key_sched else get_body_row_element_ids(schedule)

    if not fields:
        headers = build_table_headers(schedule, len(body_rows[0]) if body_rows else 0)
        body_rows, aligned_ids = _strip_body_rows(body_rows, headers, raw_ids)
        if not key_sched:
            headers, body_rows = _inject_element_ids(headers, body_rows, aligned_ids)
        return headers, body_rows, len(elements)

    headers = [field["heading"] for field in fields]

    # Use direct cell text when available and column count matches visible fields.
    if body_rows and len(body_rows[0]) == len(fields):
        body_rows, aligned_ids = _strip_body_rows(body_rows, headers, raw_ids)
        if not key_sched:
            headers, body_rows = _inject_element_ids(headers, body_rows, aligned_ids)
        return headers, body_rows, len(elements)

    # Fall back to element-based parameter lookup when GetCellText gives nothing
    # or the column count doesn't match.
    row_ids, row_lookup = build_schedule_row_lookup(schedule)
    ordered_elements = order_schedule_elements(schedule, elements, row_ids)
    if not ordered_elements:
        body_rows, aligned_ids = _strip_body_rows(body_rows, headers, raw_ids)
        if not key_sched:
            headers, body_rows = _inject_element_ids(headers, body_rows, aligned_ids)
        return headers, body_rows, len(elements)

    rows = []
    for element in ordered_elements:
        row = [get_field_text_for_element(doc, schedule, element, field, row_lookup) for field in fields]
        rows.append(row)

    # If the element-based approach also produces all-empty data, prefer the
    # raw cell-text rows even when the column count differs.
    if rows and all(all(cell == "" for cell in row) for row in rows) and body_rows:
        body_rows, aligned_ids = _strip_body_rows(body_rows, headers, raw_ids)
        if not key_sched:
            headers, body_rows = _inject_element_ids(headers, body_rows, aligned_ids)
        return headers, body_rows, len(elements)

    # Element-based path: IDs come from the ordered elements directly.
    if not key_sched:
        elem_ids = [element_id_value(e.Id) for e in ordered_elements]
        headers, rows = _inject_element_ids(headers, rows, elem_ids)
    return headers, rows, len(elements)


def auto_fit_columns(sheet):
    for column_cells in sheet.columns:
        max_length = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        column_letter = column_cells[0].column_letter
        sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 60)


def make_unique_name(base, used):
    candidate = base
    if candidate not in used:
        used.add(candidate)
        return candidate
    idx = 1
    while True:
        suffix = "_{}".format(idx)
        trimmed = candidate[: 31 - len(suffix)]
        attempt = "{}{}".format(trimmed, suffix)
        if attempt not in used:
            used.add(attempt)
            return attempt
        idx += 1


def export_to_excel(doc, category_name, category_id, param_names, file_path, ui, sheet_name=None, units_mode="project", locked_messages=None):
    """Write one category sheet. Returns True on success. If the workbook is open in
    Excel, the "close it" message is appended to locked_messages when a list is given
    (batch export), otherwise shown as an alert, and False is returned."""
    add_lib_path()
    try:
        import WWP_xlsx as openpyxl
        from WWP_xlsx import Font
    except Exception as exc:
        ui.uiUtils_alert("Excel writer is not available.\n{}".format(exc), title="Export2Ex Beta")
        return False

    def _report_locked():
        message = openpyxl.file_locked_message(file_path)
        log_message("export_to_excel locked file_path='{}'".format(file_path))
        if locked_messages is not None:
            locked_messages.append(message)
        else:
            ui.uiUtils_alert(message, title="Export2Ex Beta")
        return False

    if openpyxl.is_file_locked(file_path):
        return _report_locked()

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

    base_name = sanitize_sheet_name(sheet_name if sheet_name else _pluralize(category_name))
    sheet_name = base_name
    if sheet_name in workbook.sheetnames:
        existing = workbook[sheet_name]
        sheet_index = workbook.worksheets.index(existing)
        workbook.remove(existing)
        sheet = workbook.create_sheet(title=sheet_name, index=sheet_index)
    else:
        sheet = workbook.create_sheet(title=sheet_name)

    headers, rows, elem_count = build_category_export_rows(doc, category_id, param_names, units_mode=units_mode)
    log_message(
        "Category '{}' resolved {} elements and {} export columns".format(
            category_name, elem_count, len(headers)
        )
    )

    for col_index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=col_index, value=header)
        cell.font = Font(bold=True)
    for row_index, row in enumerate(rows, start=2):
        for col_index, value in enumerate(row, start=1):
            sheet.cell(row=row_index, column=col_index, value=value)

    if headers:
        sheet.freeze_panes = "A2"
    auto_fit_columns(sheet)

    if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
        workbook.remove(workbook["Sheet"])

    try:
        workbook.save(file_path)
    except Exception as exc:
        # Excel can open the file while the sheet is being built.
        if openpyxl.is_file_lock_error(exc):
            return _report_locked()
        raise
    return True


def show_error_report(ui, exc):
    report = "Export2Ex Beta failed.\n\n{}\n\nLog File\n{}".format(str(exc), _log_file_path())
    try:
        ui.uiUtils_alert(report, title="Export2Ex Beta")
    except Exception:
        pass


def main():
    log_message("main start")
    doc = get_active_doc()
    ui = load_uiutils()
    if doc is None:
        ui.uiUtils_alert("No active Revit document found.", title="Export2Ex Beta")
        return

    schedules = [ScheduleItem(view) for view in collect_schedules(doc)]
    categories = [CategoryItem(record) for record in collect_category_records(doc)]

    config, save_config = get_config_and_saver()
    default_dir = get_default_dir(doc)
    last_excel_path = config_get(config, CONFIG_LAST_EXCEL_PATH, "")
    init_excel_path = last_excel_path or os.path.join(default_dir, "CategoryExport.xlsx")
    last_mode = _normalize_mode(config_get(config, CONFIG_LAST_MODE, MODE_BY_CATEGORY))
    last_source_id = config_get(
        config,
        CONFIG_LAST_CATEGORY_ID if last_mode == MODE_BY_CATEGORY else CONFIG_LAST_SCHEDULE_ID,
        None,
    )
    last_category_id = _coerce_int(config_get(config, CONFIG_LAST_CATEGORY_ID, None), None)
    last_param_names = _coerce_string_list(config_get(config, CONFIG_LAST_PARAM_NAMES, []))

    def _handle_single_export(result):
        if not result:
            return
        category_id_value = _coerce_int(result.get("category_id"), None)
        if category_id_value in (None, -1):
            ui.uiUtils_alert("Select a schedule or category with a valid category.", title="Export2Ex Beta")
            return
        selected_param_names = result.get("selected_param_names") or []
        if not selected_param_names:
            ui.uiUtils_alert("Select at least one parameter to export.", title="Export2Ex Beta")
            return
        file_path = normalize_excel_output_path(_expand_path(result.get("excel_path")))
        if not file_path:
            ui.uiUtils_alert("Choose an Excel file path ending with .xlsx or .xlsm.", title="Export2Ex Beta")
            return
        category_record = None
        for item in categories:
            if item.id_value == category_id_value:
                category_record = item.record
                break
        category_name = category_record["name"] if category_record else (result.get("source_name") or "Category Export")
        sheet_name_raw = (result.get("sheet_name") or "").strip()
        if not export_to_excel(doc, category_name, DB.ElementId(Int64(category_id_value)), selected_param_names, file_path, ui, sheet_name=sheet_name_raw, units_mode=result.get("units_mode", "project")):
            return
        config.last_mode = _normalize_mode(result.get("mode"))
        config.last_schedule_id = result.get("source_id") if config.last_mode == MODE_FROM_SCHEDULE else None
        config.last_category_id = category_id_value
        config.last_param_names = list(selected_param_names)
        config.last_excel_path = file_path
        config.last_sheet_name = sheet_name_raw
        save_config()
        # Remember the file on the saved set of that name, so batch export reuses it.
        # Only the path changes; the edited parameters are saved by "Save set".
        set_name = (result.get("set_name") or "").strip()
        current_sets = read_saved_sets(doc) if set_name else {}
        existing = current_sets.get(set_name)
        if isinstance(existing, dict) and _same_saved_path(existing.get("excel_path")) != _portable_path(file_path):
            existing["excel_path"] = _portable_path(file_path)
            if not write_saved_sets(doc, current_sets):
                log_message("Could not save Excel path for set '{}': {}".format(set_name, _last_write_error))
        try:
            os.startfile(file_path)
        except Exception:
            pass
        ui.uiUtils_alert("Export complete.", title="Export2Ex Beta")

    def _open_form(set_name="", set_mode=None, set_source_id=None, set_category_id=None,
                   set_param_names=None, set_sheet_name="", set_excel_path=""):
        form_mode = _normalize_mode(set_mode) if set_mode else last_mode
        form_source_id = set_source_id if set_source_id is not None else last_source_id
        form_cat_id = _coerce_int(set_category_id, None) if set_category_id is not None else last_category_id
        form_params = _coerce_string_list(set_param_names) if set_param_names else last_param_names
        form_excel = set_excel_path or init_excel_path
        result = show_export_form(
            ui, doc, schedules, categories,
            form_excel, form_mode, form_source_id, form_cat_id, form_params,
            set_sheet_name, set_name,
        )
        if result and "batch_sets" not in result:
            _handle_single_export(result)

    def _add_callback():
        _open_form()

    def _edit_callback(sname, sdata):
        _open_form(
            set_name=sname,
            set_mode=sdata.get("mode"),
            set_source_id=sdata.get("source_id"),
            set_category_id=sdata.get("category_id"),
            set_param_names=sdata.get("param_names"),
            set_sheet_name=(sdata.get("sheet_name") or "").strip(),
            set_excel_path=(sdata.get("excel_path") or "").strip(),
        )

    saved_sets = read_saved_sets(doc)
    batch_result = _show_beta_batch_dialog(
        saved_sets, doc, ui=ui, add_callback=_add_callback, edit_callback=_edit_callback
    )

    if batch_result is None:
        return

    selected, batch_paths, batch_sheet_names = batch_result
    if not selected:
        ui.uiUtils_alert("No sets selected for batch export.", title="Export2Ex Beta")
        return

    exported = []
    failed = []
    for sname in selected:
        sdata = saved_sets.get(sname)
        if not isinstance(sdata, dict):
            failed.append("{}: set not found".format(sname))
            continue
        cat_id_val = _coerce_int(sdata.get("category_id"), None)
        if cat_id_val in (None, -1):
            failed.append("{}: no category".format(sname))
            continue
        param_names = sdata.get("param_names") or []
        if not param_names:
            failed.append("{}: no parameters configured".format(sname))
            continue
        file_path = normalize_excel_output_path(_expand_path(batch_paths.get(sname) or ""))
        if not file_path:
            failed.append("{}: no output file path".format(sname))
            continue
        sheet_name_val = (batch_sheet_names.get(sname) or "").strip()
        category_record = None
        for item in categories:
            if item.id_value == cat_id_val:
                category_record = item.record
                break
        category_name = category_record["name"] if category_record else sname
        locked_messages = []
        try:
            ok = export_to_excel(
                doc, category_name, DB.ElementId(Int64(cat_id_val)),
                param_names, file_path, ui, sheet_name=sheet_name_val,
                locked_messages=locked_messages,
            )
            if locked_messages:
                failed.append("{}: {}".format(sname, locked_messages[0]))
            elif ok:
                exported.append(sname)
            else:
                failed.append("{}: export failed".format(sname))
        except Exception as exc:
            log_exception("batch export {}".format(sname), exc)
            failed.append("{}: {}".format(sname, str(exc)))

    changed = False
    for sname, new_path in batch_paths.items():
        sdata = saved_sets.get(sname)
        if not isinstance(sdata, dict):
            continue
        if new_path and new_path != _same_saved_path(sdata.get("excel_path")):
            sdata["excel_path"] = new_path
            changed = True
    for sname, new_sheet in batch_sheet_names.items():
        sdata = saved_sets.get(sname)
        if not isinstance(sdata, dict):
            continue
        if new_sheet and new_sheet != (sdata.get("sheet_name") or "").strip():
            sdata["sheet_name"] = new_sheet
            changed = True
    if changed:
        write_saved_sets(doc, saved_sets)

    msg_parts = []
    if exported:
        msg_parts.append("Exported: {}".format(", ".join(exported)))
    if failed:
        msg_parts.append("Failed:\n{}".format("\n".join(failed)))
    ui.uiUtils_alert(
        "\n\n".join(msg_parts) if msg_parts else "Nothing exported.",
        title="Export2Ex Beta",
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log_exception("Unhandled exception in Export2Ex Beta", exc)
        try:
            show_error_report(load_uiutils(), exc)
        except Exception:
            pass
        raise
