#!python3
from __future__ import annotations

import os
import sys

from pyrevit import DB, revit


TITLE = "Import Area Key Schedule"
SKIP_OPTION = "(Skip)"
KEY_NAME_OPTION = "Key Name"


def _is_bic_value(cat_id, bic):
    if not cat_id:
        return False
    try:
        return cat_id.IntegerValue == int(bic)
    except Exception:
        return False


def choose_schedule_target(ui):
    options = ["Area Key Schedule", "Room Key Schedule"]
    selected = ui.uiUtils_select_indices(
        options,
        title=TITLE,
        prompt="Choose key schedule type to import:",
        multiselect=False,
        width=460,
        height=280,
    )
    if not selected:
        return None
    idx = selected[0]
    if idx == 1:
        return {"bic": DB.BuiltInCategory.OST_Rooms, "label": "Room"}
    return {"bic": DB.BuiltInCategory.OST_Areas, "label": "Area"}


def add_lib_path():
    script_dir = os.path.dirname(__file__)
    lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "..", "lib"))
    if lib_path not in sys.path:
        sys.path.append(lib_path)


def load_uiutils():
    add_lib_path()
    import WWP_uiUtils as ui
    return ui


def unique_view_name(doc, base_name):
    existing = set(v.Name for v in DB.FilteredElementCollector(doc).OfClass(DB.View))
    if base_name not in existing:
        return base_name
    index = 1
    while True:
        candidate = "{} ({})".format(base_name, index)
        if candidate not in existing:
            return candidate
        index += 1


def element_id_value(elem_id):
    if elem_id is None:
        return None
    if hasattr(elem_id, "IntegerValue"):
        return elem_id.IntegerValue
    if hasattr(elem_id, "Value"):
        return elem_id.Value
    try:
        return int(elem_id)
    except Exception:
        return None


def get_category_parameter_options(doc, category_bic):
    params = {}
    elements = (
        DB.FilteredElementCollector(doc)
        .OfCategory(category_bic)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    sample = elements[0] if elements else None
    if not sample:
        return [], {}

    for param in sample.Parameters:
        if not param:
            continue
        try:
            if param.IsReadOnly:
                continue
        except Exception:
            pass
        try:
            storage_type = param.StorageType
            none_type = getattr(DB.StorageType, "None")
            if storage_type == none_type:
                continue
            if storage_type == DB.StorageType.ElementId:
                continue
        except Exception:
            continue
        try:
            name = param.Definition.Name
        except Exception:
            continue
        if not name:
            continue
        param_id = param.Id
        if name in params and params[name] != param_id:
            existing_id = params.get(name)
            if existing_id is not None:
                existing_display = "{} (Id {})".format(name, element_id_value(existing_id))
                if existing_display not in params:
                    params[existing_display] = existing_id
                if name in params:
                    params.pop(name, None)
            display_name = "{} (Id {})".format(name, element_id_value(param_id))
        else:
            display_name = name
        params[display_name] = param_id

    display_names = sorted(params.keys())
    return display_names, params


def read_workbook(path, ui):
    add_lib_path()
    try:
        import openpyxl
    except Exception as exc:
        ui.uiUtils_alert("openpyxl is not available.\n{}".format(exc), title=TITLE)
        return None
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:
        ui.uiUtils_alert("Failed to open workbook.\n{}".format(exc), title=TITLE)
        return None


def _excel_column_letter(index):
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def extract_excel_data(workbook):
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return [], [], []

    max_col = 0
    for row in rows:
        if row is None:
            continue
        for idx in range(len(row) - 1, -1, -1):
            val = row[idx]
            if val not in (None, ""):
                if idx + 1 > max_col:
                    max_col = idx + 1
                break
        if len(row) > max_col:
            max_col = len(row)
    if max_col == 0:
        return [], [], []

    header_row = list(rows[0]) if rows else []
    headers = []
    column_labels = []
    for idx in range(max_col):
        raw = header_row[idx] if idx < len(header_row) else None
        header = str(raw).strip() if raw is not None else ""
        headers.append(header)
        letter = _excel_column_letter(idx + 1)
        label = letter
        if header:
            label = "{} - {}".format(letter, header)
        column_labels.append(label)

    data_rows = []
    for raw_row in rows[1:]:
        if raw_row is None:
            continue
        row = list(raw_row) + [None] * (max_col - len(raw_row))
        if all(cell in (None, "") for cell in row):
            continue
        data_rows.append(row)

    return headers, column_labels, data_rows


def format_preview_value(value):
    if value is None:
        return ""
    try:
        if hasattr(value, "isoformat"):
            return value.isoformat()
    except Exception:
        pass
    try:
        return str(value)
    except Exception:
        return ""


def build_preview_lines(headers, rows, max_rows=8):
    lines = []
    if headers:
        lines.append(" | ".join([h or "" for h in headers]))
    for row in rows[:max_rows]:
        line = " | ".join([format_preview_value(v) for v in row])
        lines.append(line)
    return lines


def build_default_selections(headers, param_display_names):
    defaults = []
    name_lookup = {name.strip().lower(): name for name in param_display_names}
    for header in headers:
        header_text = (header or "").strip()
        if not header_text:
            defaults.append(SKIP_OPTION)
            continue
        header_lower = header_text.lower()
        if header_lower in ("key", "key name", "keyname"):
            defaults.append(KEY_NAME_OPTION)
            continue
        if header_lower in name_lookup:
            defaults.append(name_lookup[header_lower])
            continue
        defaults.append(SKIP_OPTION)
    return defaults


def create_key_schedule(doc, name, category_bic, category_label):
    cat_id = DB.ElementId(category_bic)
    schedule = None
    try:
        schedule = DB.ViewSchedule.CreateKeySchedule(doc, cat_id)
    except Exception:
        schedule = None
    if schedule is None:
        try:
            schedule = DB.ViewSchedule.CreateSchedule(doc, cat_id)
        except Exception:
            schedule = None
    if schedule is None:
        return None, "Failed to create {} key schedule.".format(category_label.lower())
    if not schedule.Definition or not schedule.Definition.IsKeySchedule:
        return None, "Created schedule is not a key schedule."
    schedule.Name = unique_view_name(doc, name)
    return schedule, ""


def add_schedule_fields(definition, column_mappings, param_map):
    added_param_ids = {}
    for mapping in column_mappings:
        if mapping.get("option") in (SKIP_OPTION, KEY_NAME_OPTION):
            continue
        param_id = param_map.get(mapping.get("option"))
        if not param_id:
            continue
        param_id_val = element_id_value(param_id)
        if param_id_val in added_param_ids:
            continue
        try:
            field = definition.AddField(DB.ScheduleFieldType.Instance, param_id)
            header = mapping.get("header") or ""
            if header:
                try:
                    field.ColumnHeading = header
                except Exception:
                    pass
            added_param_ids[param_id_val] = field
        except Exception:
            continue
    return added_param_ids


def get_key_schedules(doc, category_bic):
    schedules = []
    for view in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule):
        try:
            if not view.Definition or not view.Definition.IsKeySchedule:
                continue
            cat_id = view.Definition.CategoryId
            if not _is_bic_value(cat_id, category_bic):
                continue
            schedules.append(view)
        except Exception:
            continue
    return schedules


def _normalize_name(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def schedule_field_names_with_ids(schedule):
    names = {}
    try:
        definition = schedule.Definition
        field_order = list(definition.GetFieldOrder())
    except Exception:
        field_order = []
    for field_id in field_order:
        try:
            field = definition.GetField(field_id)
        except Exception:
            continue
        if not field:
            continue
        try:
            name = field.GetName()
        except Exception:
            name = ""
        try:
            heading = field.ColumnHeading
        except Exception:
            heading = ""
        for label in (name, heading):
            key = _normalize_name(label)
            if key and key not in names:
                names[key] = field_id
    return names


def build_field_index_map(definition):
    field_order = list(definition.GetFieldOrder()) if definition else []
    index_map = {}
    for idx, field_id in enumerate(field_order):
        try:
            field = definition.GetField(field_id)
        except Exception:
            continue
        if not field:
            continue
        try:
            param_id = field.ParameterId
        except Exception:
            param_id = None
        if param_id:
            index_map[element_id_value(param_id)] = idx
    return index_map


def format_cell_value(value):
    if value is None:
        return ""
    try:
        if hasattr(value, "isoformat"):
            return value.isoformat()
    except Exception:
        pass
    try:
        return str(value)
    except Exception:
        return ""


def _as_int(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return 1 if value else 0
    if isinstance(value, (int, float)):
        try:
            return int(value)
        except Exception:
            return None
    try:
        return int(str(value).strip())
    except Exception:
        return None


def _as_yes_no_int(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return 1 if value else 0
    if isinstance(value, (int, float)):
        try:
            return 1 if int(value) != 0 else 0
        except Exception:
            return None
    text = str(value).strip().lower()
    if text in ("yes", "y", "true", "t", "1", "on", "checked", "x"):
        return 1
    if text in ("no", "n", "false", "f", "0", "off", "unchecked"):
        return 0
    return None


def _is_yes_no_parameter(param):
    try:
        definition = param.Definition
    except Exception:
        definition = None
    if not definition:
        return False
    try:
        if hasattr(definition, "GetDataType") and hasattr(DB, "SpecTypeId"):
            spec = definition.GetDataType()
            bool_yesno = getattr(getattr(DB.SpecTypeId, "Boolean", None), "YesNo", None)
            if bool_yesno and spec == bool_yesno:
                return True
    except Exception:
        pass
    try:
        parameter_type = definition.ParameterType
        if hasattr(DB, "ParameterType") and parameter_type == DB.ParameterType.YesNo:
            return True
    except Exception:
        pass
    return False


def _as_float(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        try:
            return float(value)
        except Exception:
            return None
    try:
        return float(str(value).strip())
    except Exception:
        return None


def set_parameter_value(param, value):
    if param is None:
        return False
    storage = None
    try:
        storage = param.StorageType
    except Exception:
        storage = None
    text = format_cell_value(value)
    try:
        if storage == DB.StorageType.String:
            return param.Set(text)
        if storage == DB.StorageType.Integer:
            if _is_yes_no_parameter(param):
                # Yes/No parameters must be persisted as Revit booleans (0/1), never text.
                yn = _as_yes_no_int(value)
                if yn is None:
                    return False
                return param.Set(yn)
            num = _as_int(value)
            if num is None:
                return param.SetValueString(text)
            return param.Set(num)
        if storage == DB.StorageType.Double:
            num = _as_float(value)
            if num is None:
                return param.SetValueString(text)
            try:
                return param.Set(num)
            except Exception:
                return param.SetValueString(text)
    except Exception:
        return False
    return False


def create_key_elements(schedule, row_count):
    table = schedule.GetTableData()
    body = table.GetSectionData(DB.SectionType.Body)
    created_ids = []
    existing_ids = set(
        element_id_value(e.Id)
        for e in DB.FilteredElementCollector(schedule.Document, schedule.Id).ToElements()
    )
    for _ in range(row_count):
        try:
            body.InsertRow(body.LastRowNumber)
        except Exception:
            body.InsertRow(body.NumberOfRows)
        elements = DB.FilteredElementCollector(schedule.Document, schedule.Id).ToElements()
        new_elem = None
        for elem in elements:
            elem_id_val = element_id_value(elem.Id)
            if elem_id_val not in existing_ids:
                new_elem = elem
                break
        if new_elem is None and elements:
            new_elem = elements[-1]
        if new_elem is not None:
            created_ids.append(element_id_value(new_elem.Id))
            existing_ids.add(element_id_value(new_elem.Id))
    return created_ids


def main():
    doc = revit.doc
    ui = load_uiutils()

    target = choose_schedule_target(ui)
    if not target:
        return
    category_bic = target["bic"]
    category_label = target["label"]

    param_names, param_map = get_category_parameter_options(doc, category_bic)
    if not param_names:
        ui.uiUtils_alert(
            "No {} instances found to read parameters from.".format(category_label),
            title=TITLE,
        )
        return

    parameter_options = [SKIP_OPTION, KEY_NAME_OPTION] + param_names

    state = {
        "file_path": "",
        "headers": [],
        "column_labels": [],
        "rows": [],
        "preview": [],
        "defaults": [],
    }

    while True:
        result = ui.uiUtils_area_keyplan_import(
            title=TITLE,
            file_path=state["file_path"],
            column_names=state["column_labels"],
            preview_lines=state["preview"],
            parameter_options=parameter_options,
            default_selections=state["defaults"],
            width=980,
            height=720,
        )
        if result is None:
            return

        file_path = result.get("file_path") or ""
        state["file_path"] = file_path

        if result.get("load_requested"):
            if not file_path:
                ui.uiUtils_alert("Please select an Excel file.", title=TITLE)
                continue
            workbook = read_workbook(file_path, ui)
            if workbook is None:
                continue
            headers, column_labels, rows = extract_excel_data(workbook)
            if not column_labels:
                ui.uiUtils_alert("No data found in the Excel file.", title=TITLE)
                continue
            preview = build_preview_lines(headers, rows)
            defaults = build_default_selections(headers, param_names)
            state.update(
                {
                    "headers": headers,
                    "column_labels": column_labels,
                    "rows": rows,
                    "preview": preview,
                    "defaults": defaults,
                }
            )
            continue

        if not state["column_labels"]:
            ui.uiUtils_alert("Load an Excel file first.", title=TITLE)
            continue

        selections = result.get("selected_options") or []
        if not selections or len(selections) != len(state["column_labels"]):
            ui.uiUtils_alert("Please review column mappings.", title=TITLE)
            continue

        column_mappings = []
        for idx, label in enumerate(state["column_labels"]):
            header = state["headers"][idx] if idx < len(state["headers"]) else ""
            option = selections[idx] if idx < len(selections) else SKIP_OPTION
            column_mappings.append(
                {
                    "index": idx,
                    "label": label,
                    "header": header,
                    "option": option,
                }
            )

        key_mapped = any(mapping["option"] == KEY_NAME_OPTION for mapping in column_mappings)
        if not key_mapped:
            ui.uiUtils_alert("Please map one column to 'Key Name'.", title=TITLE)
            continue

        default_name = "{} Key Schedule - Imported".format(category_label)
        if state.get("file_path"):
            try:
                base = os.path.splitext(os.path.basename(state["file_path"]))[0]
                if base:
                    default_name = base
            except Exception:
                pass
        schedule_name = ui.uiUtils_prompt_text(
            title=TITLE,
            prompt="Enter the new {} Key Schedule name:".format(category_label),
            default_value=default_name,
            ok_text="Create",
            cancel_text="Cancel",
            width=520,
            height=220,
        )
        if not schedule_name:
            return

        created = 0
        skipped = 0
        errors = []

        with revit.Transaction(TITLE):
            # Check for overlapping fields in existing area key schedules
            mapped_fields = set()
            for mapping in column_mappings:
                option = mapping.get("option")
                if option in (None, "", SKIP_OPTION, KEY_NAME_OPTION):
                    continue
                cleaned = option
                if cleaned.endswith(")") and " (Id " in cleaned:
                    cleaned = cleaned.split(" (Id ", 1)[0].strip()
                mapped_fields.add(cleaned.strip().lower())
            overlaps = []
            for sched in get_key_schedules(doc, category_bic):
                field_map = schedule_field_names_with_ids(sched)
                shared = sorted(set(field_map.keys()) & mapped_fields)
                if shared:
                    overlaps.append((sched, shared, field_map))
            # Delete existing key schedules for the selected category before creating a new one
            for sched in get_key_schedules(doc, category_bic):
                try:
                    doc.Delete(sched.Id)
                except Exception:
                    pass

            schedule, err = create_key_schedule(doc, schedule_name, category_bic, category_label)
            if schedule is None:
                ui.uiUtils_alert(err or "Failed to create key schedule.", title=TITLE)
                return

            definition = schedule.Definition
            add_schedule_fields(definition, column_mappings, param_map)

            key_param_name = ""
            try:
                key_param_name = schedule.KeyScheduleParameterName
            except Exception:
                key_param_name = "Key Name"

            row_ids = create_key_elements(schedule, len(state["rows"]))
            elements_by_id = {
                element_id_value(e.Id): e
                for e in DB.FilteredElementCollector(doc, schedule.Id).ToElements()
            }

            for idx, row in enumerate(state["rows"]):
                elem_id_val = row_ids[idx] if idx < len(row_ids) else None
                elem = elements_by_id.get(elem_id_val)
                if elem is None:
                    skipped += 1
                    continue

                for mapping in column_mappings:
                    option = mapping.get("option")
                    if option in (None, "", SKIP_OPTION):
                        continue
                    col_index = mapping.get("index", 0)
                    value = row[col_index] if col_index < len(row) else None

                    if option == KEY_NAME_OPTION:
                        param = elem.LookupParameter(key_param_name) or elem.LookupParameter("Key Name")
                        if not param:
                            errors.append("Row {} (Key Name): parameter not found".format(idx + 1))
                            continue
                        if not set_parameter_value(param, value):
                            errors.append("Row {} (Key Name): failed to set value".format(idx + 1))
                        continue

                    param = None
                    param_id = param_map.get(option)
                    if param_id:
                        try:
                            param = elem.get_Parameter(param_id)
                        except Exception:
                            param = None
                    if not param:
                        param_name = option
                        if param_name.endswith(")") and " (Id " in param_name:
                            param_name = param_name.split(" (Id ", 1)[0].strip()
                        param = elem.LookupParameter(param_name)
                    if not param:
                        errors.append("Row {} ({}): parameter not found".format(idx + 1, option))
                        continue
                    if not set_parameter_value(param, value):
                        errors.append("Row {} ({}): failed to set value".format(idx + 1, option))

                created += 1

        summary = "Rows created: {}\nRows skipped: {}".format(created, skipped)
        if errors:
            summary += "\n\nErrors:\n" + "\n".join(errors[:10])
            if len(errors) > 10:
                summary += "\n... ({} more)".format(len(errors) - 10)
        ui.uiUtils_alert(summary, title=TITLE)
        return


if __name__ == "__main__":
    main()
