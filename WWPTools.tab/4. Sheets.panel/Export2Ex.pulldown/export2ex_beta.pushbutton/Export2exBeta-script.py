#! python3

import json
import os
import re
import sys
import time
import traceback

import clr
from System import String
from System.Collections.Generic import List
from System.IO import File

from pyrevit import DB
from WWP_settings import get_tool_settings
from WWP_versioning import apply_window_title


CONFIG_LAST_EXCEL_PATH = "last_excel_path"
CONFIG_LAST_SCHEDULE_IDS = "last_schedule_ids"
LOG_FILE_NAME = "Export2ExBeta.log"
ALLOWED_EXCEL_EXTENSIONS = (".xlsx", ".xlsm")


def sanitize_sheet_name(name):
    safe = re.sub(r"[:\\/?*\[\]]", "_", (name or "").strip())
    return (safe or "Schedule")[:31]


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
        return elem_id.IntegerValue
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


def collect_schedules(doc):
    schedules = []
    for view in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule):
        if view.IsTemplate or view.IsTitleblockRevisionSchedule:
            continue
        schedules.append(view)
    schedules.sort(key=lambda item: item.Name)
    return schedules


class ScheduleItem(object):
    def __init__(self, view):
        self.view = view
        self.display_name = "{} [id:{}]".format(view.Name, element_id_value(view.Id)).replace("_", "__")


def _to_net_list(values):
    result = List[String]()
    for value in values:
        result.Add("" if value is None else str(value))
    return result


def show_export_form(ui, items, prechecked_indices, init_excel_path):
    clr.AddReference("PresentationFramework")
    clr.AddReference("PresentationCore")
    clr.AddReference("WindowsBase")
    from System.IO import StringReader
    from System import Uri
    from System.Windows.Markup import XamlReader
    from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage
    from System.Xml import XmlReader

    xaml_path = os.path.join(os.path.dirname(__file__), "ExportSchedulesDialog.xaml")
    reader = XmlReader.Create(StringReader(File.ReadAllText(xaml_path)))
    window = XamlReader.Load(reader)
    apply_window_title(window, "Export2Ex Beta")

    search_box = window.FindName("SearchBox")
    schedule_list = window.FindName("ScheduleList")
    excel_path = window.FindName("ExcelPath")
    browse_excel = window.FindName("BrowseExcel")
    ok_button = window.FindName("OkButton")
    cancel_button = window.FindName("CancelButton")
    logo_image = window.FindName("LogoImage")

    schedule_list.ItemsSource = _to_net_list(items)
    excel_path.Text = init_excel_path or ""

    selected_names = set()
    for idx in prechecked_indices or []:
        if 0 <= idx < len(items):
            selected_names.add(items[idx])

    def _apply_selection():
        schedule_list.SelectedItems.Clear()
        for item in schedule_list.Items:
            if str(item) in selected_names:
                schedule_list.SelectedItems.Add(item)

    def _filter_list(_sender=None, _args=None):
        text = (search_box.Text or "").strip().lower()
        filtered = items if not text else [item for item in items if text in item.lower()]
        schedule_list.ItemsSource = _to_net_list(filtered)
        _apply_selection()

    def _selection_changed(_sender, _args):
        selected_names.clear()
        for item in schedule_list.SelectedItems:
            selected_names.add(str(item))

    def _browse_excel(_sender, _args):
        current = excel_path.Text or ""
        file_path = ui.uiUtils_save_file_dialog(
            title="Export Schedules to Excel",
            filter_text="Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            default_extension="xlsx",
            initial_directory=os.path.dirname(current) if current else "",
            file_name=os.path.basename(current) if current else "Schedules.xlsx",
        )
        if file_path:
            excel_path.Text = file_path

    def _ok(_sender, _args):
        window.DialogResult = True
        window.Close()

    def _cancel(_sender, _args):
        window.DialogResult = False
        window.Close()

    try:
        logo_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "lib", "WWPtools-logo.png"))
        if logo_image is not None and os.path.isfile(logo_path):
            bitmap = BitmapImage()
            bitmap.BeginInit()
            bitmap.UriSource = Uri(logo_path)
            bitmap.CacheOption = BitmapCacheOption.OnLoad
            bitmap.EndInit()
            logo_image.Source = bitmap
    except Exception:
        pass

    search_box.TextChanged += _filter_list
    schedule_list.SelectionChanged += _selection_changed
    browse_excel.Click += _browse_excel
    ok_button.Click += _ok
    cancel_button.Click += _cancel
    _apply_selection()

    if not window.ShowDialog():
        return None

    return {
        "selected_indices": [idx for idx, item in enumerate(items) if item in selected_names],
        "excel_path": excel_path.Text or "",
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


def build_schedule_export_rows(doc, schedule):
    fields = get_visible_schedule_fields(schedule)
    elements = collect_schedule_elements(schedule)
    row_ids, row_lookup = build_schedule_row_lookup(schedule)
    if not fields:
        body_rows = build_schedule_body_rows(schedule)
        headers = build_table_headers(schedule, len(body_rows[0]) if body_rows else 0)
        return headers, body_rows, len(elements)

    headers = [field["heading"] for field in fields]
    ordered_elements = order_schedule_elements(schedule, elements, row_ids)
    if not ordered_elements:
        body_rows = build_schedule_body_rows(schedule)
        return headers, body_rows, len(elements)

    rows = []
    for element in ordered_elements:
        row = [get_field_text_for_element(doc, schedule, element, field, row_lookup) for field in fields]
        rows.append(row)
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


def export_to_excel(doc, schedules, file_path, ui):
    add_lib_path()
    try:
        import openpyxl
        from openpyxl.styles import Font
    except Exception as exc:
        ui.uiUtils_alert("openpyxl is not available.\n{}".format(exc), title="Export2Ex Beta")
        return False

    if os.path.exists(file_path):
        load_kwargs = {}
        if os.path.splitext(file_path)[1].lower() == ".xlsm":
            load_kwargs["keep_vba"] = True
        workbook = openpyxl.load_workbook(file_path, **load_kwargs)
    else:
        workbook = openpyxl.Workbook()

    used_sheet_names = set()
    for schedule in schedules:
        log_message("Exporting schedule '{}' ({})".format(schedule.Name, element_id_value(schedule.Id)))
        base_name = sanitize_sheet_name(schedule.Name)
        sheet_name = make_unique_name(base_name, used_sheet_names)
        if sheet_name in workbook.sheetnames:
            existing = workbook[sheet_name]
            sheet_index = workbook.worksheets.index(existing)
            workbook.remove(existing)
            sheet = workbook.create_sheet(title=sheet_name, index=sheet_index)
        else:
            sheet = workbook.create_sheet(title=sheet_name)

        headers, rows, elem_count = build_schedule_export_rows(doc, schedule)
        log_message(
            "Schedule '{}' resolved {} elements and {} visible fields".format(
                schedule.Name, elem_count, len(headers)
            )
        )

        for col_index, header in enumerate(headers, start=1):
            cell = sheet.cell(row=1, column=col_index, value=header)
            cell.font = Font(bold=True)
        for row_index, row in enumerate(rows, start=2):
            for col_index, value in enumerate(row, start=1):
                sheet.cell(row=row_index, column=col_index, value="" if value is None else str(value))

        if headers:
            sheet.freeze_panes = "A2"
        auto_fit_columns(sheet)

    if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
        workbook.remove(workbook["Sheet"])

    workbook.save(file_path)
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

    schedules = collect_schedules(doc)
    if not schedules:
        ui.uiUtils_alert("No schedules found.", title="Export2Ex Beta")
        return

    config, save_config = get_config_and_saver()
    items = [ScheduleItem(view) for view in schedules]
    last_ids = config_get(config, CONFIG_LAST_SCHEDULE_IDS, [])
    try:
        prechecked_ids = set(int(value) for value in last_ids)
    except Exception:
        prechecked_ids = set()
    prechecked_indices = [
        idx for idx, item in enumerate(items)
        if element_id_value(item.view.Id) in prechecked_ids
    ]

    default_dir = get_default_dir(doc)
    last_excel_path = config_get(config, CONFIG_LAST_EXCEL_PATH, "")
    init_excel_path = last_excel_path or os.path.join(default_dir, "Schedules.xlsx")
    result = show_export_form(
        ui,
        [item.display_name for item in items],
        prechecked_indices,
        init_excel_path,
    )
    if not result:
        return

    selected_indices = result.get("selected_indices") or []
    if not selected_indices:
        ui.uiUtils_alert("Select at least one schedule.", title="Export2Ex Beta")
        return

    file_path = normalize_excel_output_path(result.get("excel_path"))
    if not file_path:
        ui.uiUtils_alert(
            "Choose an Excel file path ending with .xlsx or .xlsm.",
            title="Export2Ex Beta",
        )
        return

    selected_views = [items[idx].view for idx in selected_indices]
    if not export_to_excel(doc, selected_views, file_path, ui):
        return

    config.last_schedule_ids = [element_id_value(view.Id) for view in selected_views]
    config.last_excel_path = file_path
    save_config()

    try:
        os.startfile(file_path)
    except Exception:
        pass
    ui.uiUtils_alert("Export complete.", title="Export2Ex Beta")


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
