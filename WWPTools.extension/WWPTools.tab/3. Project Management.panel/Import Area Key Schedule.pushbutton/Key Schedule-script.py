#!python3
from __future__ import annotations

import os
import sys

from pyrevit import DB, revit, script


TITLE = "Import Area Key Schedule"
SKIP_OPTION = "(Skip)"
KEY_NAME_OPTION = "Key Name"
TARGET_OPTIONS = ["Area Key Schedule", "Room Key Schedule"]
DEFAULT_SCHEDULE_SUFFIX = "Key Schedule - Imported"


def _get_config():
    try:
        return script.get_config()
    except Exception:
        return None


def _safe_config_get(config, key, default=None):
    if config is None:
        return default
    try:
        return getattr(config, key, default)
    except Exception:
        return default


def _safe_config_set(config, key, value):
    if config is None:
        return
    try:
        setattr(config, key, value)
    except Exception:
        pass


def _save_config():
    try:
        script.save_config()
    except Exception:
        pass


def _header_signature(headers):
    return "|".join([_normalize_name(h) for h in headers or []])


def _sanitize_selections(selections, column_count, parameter_options):
    if not selections or len(selections) != column_count:
        return None
    valid = set(parameter_options)
    cleaned = []
    for item in selections:
        if item in valid:
            cleaned.append(item)
        else:
            cleaned.append(SKIP_OPTION)
    return cleaned


def _default_schedule_name(file_path, category_label):
    default_name = "{} {}".format(category_label, DEFAULT_SCHEDULE_SUFFIX)
    if file_path:
        try:
            base = os.path.splitext(os.path.basename(file_path))[0]
            if base:
                return base
        except Exception:
            pass
    return default_name


def _is_bic_value(cat_id, bic):
    if not cat_id:
        return False
    try:
        return cat_id.IntegerValue == int(bic)
    except Exception:
        return False


def resolve_schedule_target(selected_target_type):
    if _normalize_name(selected_target_type) == _normalize_name(TARGET_OPTIONS[1]):
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
    name_lookup = {}
    for name in param_display_names:
        raw = (name or "").strip()
        if not raw:
            continue
        key = _normalize_name(raw)
        if key and key not in name_lookup:
            name_lookup[key] = name
        # Also support display names like "Param Name (Id 123)".
        base = raw
        if base.endswith(")") and " (id " in _normalize_name(base):
            base = base.rsplit(" (", 1)[0].strip()
            base_key = _normalize_name(base)
            if base_key and base_key not in name_lookup:
                name_lookup[base_key] = name
    for header in headers:
        header_text = (header or "").strip()
        if not header_text:
            defaults.append(SKIP_OPTION)
            continue
        header_lower = _normalize_name(header_text)
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


def find_key_schedule_by_name(doc, category_bic, schedule_name):
    target = _normalize_name(schedule_name)
    if not target:
        return None
    for sched in get_key_schedules(doc, category_bic):
        try:
            if _normalize_name(sched.Name) == target:
                return sched
        except Exception:
            continue
    return None


def _normalize_name(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def get_schedule_key_parameter_names(schedule):
    names = []
    try:
        key_name = (schedule.KeyScheduleParameterName or "").strip()
        if key_name:
            names.append(key_name)
    except Exception:
        pass
    if "Key Name" not in names:
        names.append("Key Name")
    return names


def get_schedule_key_name_value(elem, key_param_names):
    for param_name in key_param_names:
        param = elem.LookupParameter(param_name)
        if not param:
            continue
        try:
            value = param.AsString()
            if value:
                return value
        except Exception:
            pass
        try:
            value = param.AsValueString()
            if value:
                return value
        except Exception:
            pass
    return ""


def collect_schedule_key_elements(schedule):
    elements = []
    try:
        elements = list(DB.FilteredElementCollector(schedule.Document, schedule.Id).ToElements())
    except Exception:
        elements = []
    return [elem for elem in elements if elem is not None]


def build_schedule_key_element_map(schedule, key_param_names):
    key_map = {}
    duplicates = set()
    for elem in collect_schedule_key_elements(schedule):
        key_value = get_schedule_key_name_value(elem, key_param_names)
        key_norm = _normalize_name(key_value)
        if not key_norm:
            continue
        if key_norm in key_map:
            duplicates.add(key_norm)
            continue
        key_map[key_norm] = elem
    return key_map, duplicates


def build_key_usage_counts(doc, category_bic, host_key_param_name, key_ids):
    usage = dict((key_id, 0) for key_id in key_ids)
    if not host_key_param_name or not key_ids:
        return usage
    for elem in (
        DB.FilteredElementCollector(doc)
        .OfCategory(category_bic)
        .WhereElementIsNotElementType()
        .ToElements()
    ):
        param = elem.LookupParameter(host_key_param_name)
        if not param:
            continue
        try:
            if param.StorageType != DB.StorageType.ElementId:
                continue
        except Exception:
            continue
        try:
            ref_id = param.AsElementId()
        except Exception:
            continue
        ref_val = element_id_value(ref_id)
        if ref_val in usage:
            usage[ref_val] = usage.get(ref_val, 0) + 1
    return usage


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


def create_single_key_element(schedule):
    created = create_key_elements(schedule, 1)
    if not created:
        return None
    elem_id = DB.ElementId(created[0])
    try:
        return schedule.Document.GetElement(elem_id)
    except Exception:
        return None


def main():
    doc = revit.doc
    ui = load_uiutils()
    config = _get_config()
    selected_target_type = _safe_config_get(config, "area_key_import_target", "") or TARGET_OPTIONS[0]
    target = resolve_schedule_target(selected_target_type)
    category_bic = target["bic"]
    category_label = target["label"]
    category_key = _normalize_name(category_label)

    param_names, param_map = get_category_parameter_options(doc, category_bic)
    parameter_options = [SKIP_OPTION, KEY_NAME_OPTION] + param_names

    last_file_path = _safe_config_get(config, "area_key_import_file_path", "") or ""

    state = {
        "file_path": last_file_path,
        "headers": [],
        "column_labels": [],
        "rows": [],
        "defaults": [],
        "schedule_name": "",
        "selected_target_type": selected_target_type,
    }

    # Preload the most recently used workbook for this target type.
    if state["file_path"] and os.path.exists(state["file_path"]):
        workbook = read_workbook(state["file_path"], ui)
        if workbook is not None:
            headers, column_labels, rows = extract_excel_data(workbook)
            if column_labels:
                signature = _header_signature(headers)
                saved_signature = _safe_config_get(config, "area_key_import_headers_signature", "") or ""
                saved_target = _safe_config_get(config, "area_key_import_target", "") or ""
                saved_selections = _safe_config_get(config, "area_key_import_selected_options", None)
                defaults = None
                if _normalize_name(saved_target) == category_key and saved_signature == signature:
                    defaults = _sanitize_selections(saved_selections, len(column_labels), parameter_options)
                if defaults is None:
                    defaults = build_default_selections(headers, param_names)
                state.update(
                    {
                        "headers": headers,
                        "column_labels": column_labels,
                        "rows": rows,
                        "defaults": defaults,
                    }
                )

    if not state["schedule_name"]:
        state["schedule_name"] = _safe_config_get(
            config,
            "area_key_import_schedule_name_{}".format(category_key),
            "",
        ) or _default_schedule_name(state["file_path"], category_label)

    while True:
        result = ui.uiUtils_area_keyplan_import(
            title=TITLE,
            file_path=state["file_path"],
            column_names=state["column_labels"],
            target_types=TARGET_OPTIONS,
            selected_target_type=state.get("selected_target_type", TARGET_OPTIONS[0]),
            parameter_options=parameter_options,
            default_selections=state["defaults"],
            width=980,
            height=720,
        )
        if result is None:
            return

        selected_target_type = result.get("selected_target_type") or state.get("selected_target_type", TARGET_OPTIONS[0])
        next_target = resolve_schedule_target(selected_target_type)
        next_category_bic = next_target["bic"]
        next_category_label = next_target["label"]
        next_category_key = _normalize_name(next_category_label)
        target_changed = next_category_bic != category_bic
        if target_changed:
            category_bic = next_category_bic
            category_label = next_category_label
            category_key = next_category_key
            state["selected_target_type"] = selected_target_type
            param_names, param_map = get_category_parameter_options(doc, category_bic)
            parameter_options = [SKIP_OPTION, KEY_NAME_OPTION] + param_names
            state["defaults"] = []
            state["schedule_name"] = _safe_config_get(
                config,
                "area_key_import_schedule_name_{}".format(category_key),
                "",
            ) or _default_schedule_name(state.get("file_path", ""), category_label)

        file_path = result.get("file_path") or ""
        state["file_path"] = file_path

        if result.get("load_requested"):
            state["selected_target_type"] = selected_target_type
            if not file_path:
                ui.uiUtils_alert("Please select an Excel file.", title=TITLE)
                continue
            if not param_names:
                ui.uiUtils_alert(
                    "No {} instances found to read parameters from.".format(category_label),
                    title=TITLE,
                )
                continue
            workbook = read_workbook(file_path, ui)
            if workbook is None:
                continue
            headers, column_labels, rows = extract_excel_data(workbook)
            if not column_labels:
                ui.uiUtils_alert("No data found in the Excel file.", title=TITLE)
                continue
            signature = _header_signature(headers)
            saved_signature = _safe_config_get(config, "area_key_import_headers_signature", "") or ""
            saved_target = _safe_config_get(config, "area_key_import_target", "") or ""
            saved_selections = _safe_config_get(config, "area_key_import_selected_options", None)
            defaults = None
            if _normalize_name(saved_target) == category_key and saved_signature == signature:
                defaults = _sanitize_selections(saved_selections, len(column_labels), parameter_options)
            if defaults is None:
                defaults = build_default_selections(headers, param_names)
            state.update(
                {
                    "headers": headers,
                    "column_labels": column_labels,
                    "rows": rows,
                    "defaults": defaults,
                }
            )
            saved_name = _safe_config_get(
                config,
                "area_key_import_schedule_name_{}".format(category_key),
                "",
            ) or ""
            state["schedule_name"] = saved_name or _default_schedule_name(state["file_path"], category_label)
            continue

        if target_changed:
            # Re-open once to refresh mapping options for the newly selected target type.
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

        schedule_name = (state.get("schedule_name") or "").strip()
        if not schedule_name:
            schedule_name = _default_schedule_name(state.get("file_path", ""), category_label)
        if not schedule_name:
            ui.uiUtils_alert("Unable to determine schedule name.", title=TITLE)
            continue

        key_column_indices = [
            m.get("index", 0)
            for m in column_mappings
            if m.get("option") == KEY_NAME_OPTION
        ]
        key_column_index = key_column_indices[0] if key_column_indices else None
        if key_column_index is None:
            ui.uiUtils_alert("Please map one column to 'Key Name'.", title=TITLE)
            continue

        created = 0
        updated = 0
        deleted = 0
        kept_in_use = 0
        skipped = 0
        errors = []
        warnings = []

        with revit.Transaction(TITLE):
            schedule = find_key_schedule_by_name(doc, category_bic, schedule_name)
            if schedule is None:
                schedule, err = create_key_schedule(doc, schedule_name, category_bic, category_label)
                if schedule is None:
                    ui.uiUtils_alert(err or "Failed to create key schedule.", title=TITLE)
                    return

            definition = schedule.Definition
            add_schedule_fields(definition, column_mappings, param_map)

            key_param_names = get_schedule_key_parameter_names(schedule)
            key_param_name = key_param_names[0] if key_param_names else "Key Name"
            existing_key_map, duplicate_existing_keys = build_schedule_key_element_map(schedule, key_param_names)
            if duplicate_existing_keys:
                warnings.append(
                    "Existing schedule has duplicate key names. Only the first match was updated for: {}".format(
                        ", ".join(sorted(list(duplicate_existing_keys))[:10])
                    )
                )
            matched_existing_keys = set()
            seen_excel_keys = set()

            for idx, row in enumerate(state["rows"]):
                key_value = row[key_column_index] if key_column_index < len(row) else None
                key_text = format_cell_value(key_value).strip()
                key_norm = _normalize_name(key_text)
                if not key_norm:
                    errors.append("Row {}: key name is blank".format(idx + 1))
                    skipped += 1
                    continue
                if key_norm in seen_excel_keys:
                    errors.append("Row {}: duplicate key name '{}' in Excel".format(idx + 1, key_text))
                    skipped += 1
                    continue
                seen_excel_keys.add(key_norm)

                elem = existing_key_map.get(key_norm)
                is_new = False
                if elem is not None:
                    matched_existing_keys.add(key_norm)
                else:
                    elem = create_single_key_element(schedule)
                    is_new = True
                if elem is None:
                    errors.append("Row {} ({}): failed to create or find key row".format(idx + 1, key_text))
                    skipped += 1
                    continue

                for mapping in column_mappings:
                    option = mapping.get("option")
                    if option in (None, "", SKIP_OPTION):
                        continue
                    col_index = mapping.get("index", 0)
                    value = row[col_index] if col_index < len(row) else None

                    if option == KEY_NAME_OPTION:
                        param = None
                        for p_name in key_param_names:
                            param = elem.LookupParameter(p_name)
                            if param:
                                break
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

                if is_new:
                    created += 1
                else:
                    updated += 1

            to_remove_keys = [
                key_norm
                for key_norm in existing_key_map.keys()
                if key_norm not in matched_existing_keys
            ]
            key_ids_to_check = [
                element_id_value(existing_key_map[k].Id)
                for k in to_remove_keys
                if existing_key_map.get(k) is not None
            ]
            usage_counts = build_key_usage_counts(doc, category_bic, key_param_name, key_ids_to_check)
            for key_norm in to_remove_keys:
                elem = existing_key_map.get(key_norm)
                if elem is None:
                    continue
                elem_id_val = element_id_value(elem.Id)
                if usage_counts.get(elem_id_val, 0) > 0:
                    kept_in_use += 1
                    continue
                try:
                    doc.Delete(elem.Id)
                    deleted += 1
                except Exception:
                    errors.append("Failed to delete old key '{}'".format(key_norm))

        # Persist latest choices for the next run.
        _safe_config_set(config, "area_key_import_target", state.get("selected_target_type", TARGET_OPTIONS[0]))
        _safe_config_set(config, "area_key_import_file_path", state.get("file_path", ""))
        _safe_config_set(config, "area_key_import_headers_signature", _header_signature(state.get("headers", [])))
        _safe_config_set(config, "area_key_import_selected_options", selections)
        _safe_config_set(config, "area_key_import_schedule_name_{}".format(category_key), schedule_name)
        _save_config()

        summary = (
            "Rows created: {}\nRows updated: {}\nRows deleted: {}\n"
            "Rows kept (in use): {}\nRows skipped: {}"
        ).format(created, updated, deleted, kept_in_use, skipped)
        if warnings:
            summary += "\n\nWarnings:\n" + "\n".join(warnings[:10])
            if len(warnings) > 10:
                summary += "\n... ({} more)".format(len(warnings) - 10)
        if errors:
            summary += "\n\nErrors:\n" + "\n".join(errors[:10])
            if len(errors) > 10:
                summary += "\n... ({} more)".format(len(errors) - 10)
        ui.uiUtils_alert(summary, title=TITLE)
        return


if __name__ == "__main__":
    main()
