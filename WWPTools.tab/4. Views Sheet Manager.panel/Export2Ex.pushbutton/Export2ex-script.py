#! python3
from __future__ import annotations

import csv
import importlib
import os
import re
import shutil
import sys
import tempfile

from pyrevit import DB, revit, script



CONFIG_LAST_EXCEL_PATH = "last_excel_path"
CONFIG_LAST_CSV_DIR = "last_csv_dir"
CONFIG_LAST_SCHEDULE_IDS = "last_schedule_ids"
CONFIG_LAST_CSV_MODE = "last_csv_mode"
CONFIG_LAST_CSV_DELIM = "last_csv_delim"
CONFIG_LAST_EXPORT_MODE = "last_export_mode"


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
        if view.IsTemplate:
            continue
        try:
            if view.ViewType == DB.ViewType.Legend:
                continue
        except Exception:
            pass
        if view.IsTitleblockRevisionSchedule:
            continue
        if view.Definition and view.Definition.IsKeySchedule:
            continue
        schedules.append(view)
    schedules.sort(key=lambda v: v.Name)
    return schedules


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


def read_csv_rows(path, delimiter=","):
    for encoding in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.reader(handle, delimiter=delimiter)
                return [row for row in reader]
        except Exception:
            continue
    return []


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


def export_to_excel(doc, schedules, file_path, ui):
    add_lib_path()
    try:
        import openpyxl
    except Exception as exc:
        ui.uiUtils_alert(
            "openpyxl is not available.\n{}".format(exc),
            title="Multiple Schedules Exporter",
        )
        return False

    used_names = set()
    if os.path.exists(file_path):
        workbook = openpyxl.load_workbook(file_path)
    else:
        workbook = openpyxl.Workbook()

    existing_names = set(workbook.sheetnames)
    existing_list = list(workbook.sheetnames)
    temp_dir = tempfile.mkdtemp(prefix="wwp_schedules_")

    options = DB.ViewScheduleExportOptions()
    options.FieldDelimiter = ","
    try:
        for view in schedules:
            base_name = sanitize_sheet_name(view.Name)
            if base_name in existing_names and base_name not in used_names:
                sheet_name = base_name
            else:
                candidates = [
                    name for name in existing_list
                    if name.startswith(base_name) and name not in used_names
                ]
                if len(candidates) == 1:
                    sheet_name = candidates[0]
                else:
                    used_pool = set(existing_names)
                    used_pool.update(used_names)
                    sheet_name = make_unique_name(base_name, used_pool, max_len=31)
            used_names.add(sheet_name)
            if sheet_name in workbook.sheetnames:
                existing = workbook[sheet_name]
                workbook.remove(existing)
                sheet = workbook.create_sheet(title=sheet_name)
            else:
                sheet = workbook.create_sheet(title=sheet_name)

            temp_name = "{}.csv".format(sanitize_file_name(view.Name))
            view.Export(temp_dir, temp_name, options)
            csv_path = os.path.join(temp_dir, temp_name)
            data = normalize_table_data(read_csv_rows(csv_path))
            data = inject_element_id_column(data, view, csv_text=False)
            header_rows = get_section_row_count(view, DB.SectionType.Header)
            column_specs = get_column_specs(view)
            if column_specs is not None:
                column_specs = [None] + column_specs
            write_table_to_sheet(
                sheet,
                data,
                1,
                header_rows=header_rows,
                column_specs=column_specs,
                doc=doc,
            )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
        default_sheet = workbook["Sheet"]
        workbook.remove(default_sheet)

    workbook.save(file_path)
    return True


def export_to_csv(doc, schedules, folder, quote_all=False, delimiter=","):
    if not os.path.isdir(folder):
        os.makedirs(folder)
    options = DB.ViewScheduleExportOptions()
    options.FieldDelimiter = delimiter
    used_names = set()
    for view in schedules:
        base_name = sanitize_file_name(view.Name)
        unique_name = make_unique_name(base_name, used_names)
        file_name = "{}.csv".format(unique_name)
        view.Export(folder, file_name, options)
        csv_path = os.path.join(folder, file_name)
        rows = normalize_table_data(read_csv_rows(csv_path, delimiter=delimiter))
        rows = inject_element_id_column(rows, view, csv_text=True)
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as handle:
            if quote_all:
                writer = csv.writer(handle, delimiter=delimiter, quoting=csv.QUOTE_ALL)
            else:
                writer = csv.writer(handle, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)
            writer.writerows(rows)
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
    doc = revit.doc
    config = script.get_config()
    ui = load_uiutils()

    schedules = collect_schedules(doc)
    if not schedules:
        ui.uiUtils_alert("No schedules found.", title="Multiple Schedules Exporter")
        return

    items = [ScheduleItem(v) for v in schedules]
    last_ids = getattr(config, CONFIG_LAST_SCHEDULE_IDS, [])
    try:
        prechecked_ids = set(int(x) for x in last_ids)
    except Exception:
        prechecked_ids = set()
    prechecked_indices = [
        idx for idx, item in enumerate(items)
        if element_id_value(item.view.Id) in prechecked_ids
    ]
    default_dir = get_default_dir(doc)
    last_excel_path = getattr(config, CONFIG_LAST_EXCEL_PATH, "")
    last_csv_dir = getattr(config, CONFIG_LAST_CSV_DIR, "")
    last_csv_mode = getattr(config, CONFIG_LAST_CSV_MODE, 0)
    last_csv_delim = getattr(config, CONFIG_LAST_CSV_DELIM, ",")
    last_export_mode = getattr(config, CONFIG_LAST_EXPORT_MODE, 0)

    init_excel_path = last_excel_path or os.path.join(default_dir, "Schedules.xlsx")
    init_csv_dir = ensure_existing_dir(last_csv_dir, default_dir)
    inputs = ui.uiUtils_export_schedules_inputs(
        [item.display_name for item in items],
        title="Export Schedules",
        prompt="Select schedules to export:",
        mode_labels=("Export to Excel", "Export to CSV"),
        default_mode=last_export_mode,
        prechecked_indices=prechecked_indices,
        excel_path=init_excel_path,
        csv_folder=init_csv_dir,
        csv_delimiter=last_csv_delim,
        csv_quote_all=(last_csv_mode == 1),
        width=860,
        height=720,
    )
    if inputs is not False:
        if not inputs:
            return
        selected_indices = inputs.get("selected_indices") or []
        if not selected_indices:
            ui.uiUtils_alert("Select at least one schedule.", title="Multiple Schedules Exporter")
            return
        selected_views = [items[i].view for i in selected_indices]
        config.last_schedule_ids = [element_id_value(v.Id) for v in selected_views]
        mode = int(inputs.get("mode", 0))
        if mode == 0:
            file_path = (inputs.get("excel_path") or "").strip()
            if not file_path:
                ui.uiUtils_alert("Choose an Excel file path.", title="Multiple Schedules Exporter")
                return
            if not file_path.lower().endswith(".xlsx"):
                file_path = "{}.xlsx".format(file_path)
            success = export_to_excel(doc, selected_views, file_path, ui)
            if not success:
                return
            config.last_excel_path = file_path
            try:
                os.startfile(file_path)
            except Exception:
                pass
        else:
            folder = (inputs.get("csv_folder") or "").strip()
            if not folder:
                ui.uiUtils_alert("Choose a CSV folder.", title="Multiple Schedules Exporter")
                return
            csv_delim = inputs.get("csv_delimiter") or ","
            quote_all = bool(inputs.get("csv_quote_all"))
            export_to_csv(doc, selected_views, folder, quote_all=quote_all, delimiter=csv_delim)
            config.last_csv_dir = folder
            config.last_csv_mode = 1 if quote_all else 0
            config.last_csv_delim = csv_delim
        config.last_export_mode = mode
        script.save_config()
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
        file_path = ui.uiUtils_save_file_dialog(
            title="Export Schedules",
            filter_text="Excel Workbook (*.xlsx)|*.xlsx",
            default_extension="xlsx",
            initial_directory=init_dir,
            file_name=os.path.basename(last_excel_path) if last_excel_path else "Schedules.xlsx",
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path = "{}.xlsx".format(file_path)
        success = export_to_excel(doc, selected_views, file_path, ui)
        if not success:
            return
        config.last_excel_path = file_path
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
        folder = ui.uiUtils_select_folder_dialog(
            title="Select CSV Folder",
            initial_directory=init_dir,
        )
        if not folder:
            return
        export_to_csv(doc, selected_views, folder, quote_all=(csv_mode == 1), delimiter=csv_delim)
        config.last_csv_dir = folder
        config.last_csv_mode = csv_mode
        config.last_csv_delim = csv_delim

    script.save_config()
    ui.uiUtils_alert("Export complete.", title="Multiple Schedules Exporter")


if __name__ == "__main__":
    main()
