#!python3
from __future__ import annotations

import os
import sys

from pyrevit import DB, revit


TITLE = "Map Key Schedule By Name"
TARGET_OPTIONS = ["Area Key Schedule", "Room Key Schedule"]
PREFERRED_PROGRAM = "residential"


def add_lib_path():
    script_dir = os.path.dirname(__file__)
    lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "..", "lib"))
    if lib_path not in sys.path:
        sys.path.append(lib_path)


add_lib_path()

from WWP_settings import get_tool_settings
import WWP_uiUtils as ui


config, save_config = get_tool_settings("MapKeyScheduleByName", doc=revit.doc)


def _safe_config_get(key, default=None):
    try:
        return getattr(config, key, default)
    except Exception:
        return default


def _safe_config_set(key, value):
    try:
        setattr(config, key, value)
    except Exception:
        pass


def _save_config():
    try:
        save_config()
    except Exception:
        pass


def _normalize_name(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def _is_bic_value(cat_id, bic):
    if not cat_id:
        return False
    try:
        return cat_id.IntegerValue == int(bic)
    except Exception:
        return False


def _element_id_value(elem_id):
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


def _get_param_string(elem, param_name):
    if elem is None or not param_name:
        return ""
    try:
        param = elem.LookupParameter(param_name)
    except Exception:
        param = None
    if not param:
        return ""
    for getter_name in ("AsString", "AsValueString"):
        try:
            value = getattr(param, getter_name)()
            if value:
                return str(value).strip()
        except Exception:
            continue
    return ""


def resolve_schedule_target(selected_target_type):
    if _normalize_name(selected_target_type) == _normalize_name(TARGET_OPTIONS[1]):
        return {"bic": DB.BuiltInCategory.OST_Rooms, "label": "Room"}
    return {"bic": DB.BuiltInCategory.OST_Areas, "label": "Area"}


def get_key_schedules(doc, category_bic):
    schedules = []
    for view in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule):
        try:
            if not view.Definition or not view.Definition.IsKeySchedule:
                continue
            if not _is_bic_value(view.Definition.CategoryId, category_bic):
                continue
            schedules.append(view)
        except Exception:
            continue
    return sorted(schedules, key=lambda item: getattr(item, "Name", "").lower())


def get_schedule_key_parameter_name(schedule):
    try:
        return (schedule.KeyScheduleParameterName or "").strip()
    except Exception:
        return ""


def get_schedule_key_parameter_names(schedule):
    names = []
    key_name = get_schedule_key_parameter_name(schedule)
    if key_name:
        names.append(key_name)
    if "Key Name" not in names:
        names.append("Key Name")
    return names


def get_schedule_key_name_value(elem, key_param_names):
    for param_name in key_param_names:
        value = _get_param_string(elem, param_name)
        if value:
            return value
    return ""


def collect_schedule_key_elements(schedule):
    try:
        elements = list(DB.FilteredElementCollector(schedule.Document, schedule.Id).ToElements())
    except Exception:
        elements = []
    return [elem for elem in elements if elem is not None]


def _sort_key_candidates(candidates):
    return sorted(
        candidates,
        key=lambda item: (
            0 if _normalize_name(item.get("program")) == PREFERRED_PROGRAM else 1,
            _normalize_name(item.get("key_name")),
            _element_id_value(getattr(item.get("element"), "Id", None)) or 0,
        ),
    )


def build_schedule_name_match_map(schedule):
    key_param_names = get_schedule_key_parameter_names(schedule)
    candidates_by_name = {}

    for elem in collect_schedule_key_elements(schedule):
        name_text = _get_param_string(elem, "Name")
        name_norm = _normalize_name(name_text)
        if not name_norm:
            continue
        candidates_by_name.setdefault(name_norm, []).append(
            {
                "element": elem,
                "name": name_text,
                "key_name": get_schedule_key_name_value(elem, key_param_names),
                "program": _get_param_string(elem, "Program"),
            }
        )

    chosen_by_name = {}
    resolved_to_residential = 0
    duplicate_samples = []

    for name_norm, candidates in candidates_by_name.items():
        ordered = _sort_key_candidates(candidates)
        chosen = ordered[0]
        chosen_by_name[name_norm] = chosen

        if len(ordered) > 1:
            residential_matches = [item for item in ordered if _normalize_name(item.get("program")) == PREFERRED_PROGRAM]
            if residential_matches and chosen in residential_matches:
                resolved_to_residential += 1
            sample = "{} -> {}".format(
                chosen.get("name") or name_norm,
                ", ".join(
                    [
                        "{} [{} / {}]".format(
                            item.get("key_name") or "(blank key)",
                            item.get("program") or "(blank program)",
                            _element_id_value(getattr(item.get("element"), "Id", None)) or "?",
                        )
                        for item in ordered[:5]
                    ]
                ),
            )
            duplicate_samples.append(sample)

    return chosen_by_name, resolved_to_residential, duplicate_samples


def get_target_elements(doc, category_bic):
    selected = []
    selection_ids = revit.uidoc.Selection.GetElementIds()
    for elem_id in selection_ids:
        try:
            elem = doc.GetElement(elem_id)
        except Exception:
            elem = None
        if elem is None:
            continue
        try:
            if elem.Category and _is_bic_value(elem.Category.Id, category_bic):
                selected.append(elem)
        except Exception:
            continue
    if selected:
        return selected, "selection"

    elements = list(
        DB.FilteredElementCollector(doc)
        .OfCategory(category_bic)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    return [elem for elem in elements if elem is not None], "model"


def get_target_name(elem, category_bic):
    built_in = None
    try:
        if category_bic == DB.BuiltInCategory.OST_Rooms:
            built_in = DB.BuiltInParameter.ROOM_NAME
        elif category_bic == DB.BuiltInCategory.OST_Areas:
            built_in = DB.BuiltInParameter.AREA_NAME
    except Exception:
        built_in = None

    if built_in is not None:
        try:
            param = elem.get_Parameter(built_in)
        except Exception:
            param = None
        if param:
            for getter_name in ("AsString", "AsValueString"):
                try:
                    value = getattr(param, getter_name)()
                    if value:
                        return str(value).strip()
                except Exception:
                    continue
    return _get_param_string(elem, "Name")


def select_single_index(items, title, prompt, default_index=0, width=640, height=420):
    if not items:
        return None
    ordered_items = list(items)
    if default_index and 0 <= default_index < len(ordered_items):
        default_item = ordered_items[default_index]
        ordered_items = [default_item] + [item for idx, item in enumerate(ordered_items) if idx != default_index]
        picked = ui.uiUtils_select_indices(ordered_items, title=title, prompt=prompt, multiselect=False, width=width, height=height)
        if not picked:
            return None
        chosen_value = ordered_items[picked[0]]
        return items.index(chosen_value)
    picked = ui.uiUtils_select_indices(ordered_items, title=title, prompt=prompt, multiselect=False, width=width, height=height)
    if not picked:
        return None
    return picked[0]


def choose_target_type():
    saved_target = _safe_config_get("map_key_schedule_target", TARGET_OPTIONS[0]) or TARGET_OPTIONS[0]
    default_index = 0
    for idx, label in enumerate(TARGET_OPTIONS):
        if _normalize_name(label) == _normalize_name(saved_target):
            default_index = idx
            break
    picked_index = select_single_index(
        TARGET_OPTIONS,
        title=TITLE,
        prompt="Choose the target category:",
        default_index=default_index,
        width=420,
        height=240,
    )
    if picked_index is None:
        return None
    return TARGET_OPTIONS[picked_index]


def choose_schedule(schedules, category_key):
    schedule_names = [sched.Name for sched in schedules]
    saved_name = _safe_config_get("map_key_schedule_name_{}".format(category_key), "") or ""
    default_index = 0
    for idx, name in enumerate(schedule_names):
        if _normalize_name(name) == _normalize_name(saved_name):
            default_index = idx
            break
    picked_index = select_single_index(
        schedule_names,
        title=TITLE,
        prompt="Choose the key schedule to map from:",
        default_index=default_index,
        width=760,
        height=520,
    )
    if picked_index is None:
        return None
    return schedules[picked_index]


def build_summary(processed_scope, category_label, schedule_name, mapped_count, already_count, unmatched_names, missing_param_count, duplicate_samples, resolved_to_residential):
    lines = [
        "Target: {}s ({})".format(category_label, processed_scope),
        "Key schedule: {}".format(schedule_name),
        "",
        "Mapped: {}".format(mapped_count),
        "Already matched: {}".format(already_count),
        "Missing key parameter: {}".format(missing_param_count),
        "Unmatched names: {}".format(len(unmatched_names)),
        "Duplicate schedule names resolved to Residential: {}".format(resolved_to_residential),
    ]
    if unmatched_names:
        lines.append("")
        lines.append("Top unmatched names:")
        for name in unmatched_names[:10]:
            lines.append("- {}".format(name))
        if len(unmatched_names) > 10:
            lines.append("- ... ({} more)".format(len(unmatched_names) - 10))
    if duplicate_samples:
        lines.append("")
        lines.append("Duplicate schedule name matches:")
        for sample in duplicate_samples[:10]:
            lines.append("- {}".format(sample))
        if len(duplicate_samples) > 10:
            lines.append("- ... ({} more)".format(len(duplicate_samples) - 10))
    return "\n".join(lines)


def main():
    doc = revit.doc

    selected_target_type = choose_target_type()
    if not selected_target_type:
        return

    target = resolve_schedule_target(selected_target_type)
    category_bic = target["bic"]
    category_label = target["label"]
    category_key = _normalize_name(category_label)

    schedules = get_key_schedules(doc, category_bic)
    if not schedules:
        ui.uiUtils_alert("No {} key schedules were found in this model.".format(category_label.lower()), title=TITLE)
        return

    schedule = choose_schedule(schedules, category_key)
    if schedule is None:
        return

    host_key_param_name = get_schedule_key_parameter_name(schedule)
    if not host_key_param_name:
        ui.uiUtils_alert("The selected key schedule does not expose a host key parameter name.", title=TITLE)
        return

    chosen_by_name, resolved_to_residential, duplicate_samples = build_schedule_name_match_map(schedule)
    if not chosen_by_name:
        ui.uiUtils_alert("No key schedule rows with a usable 'Name' value were found.", title=TITLE)
        return

    elements, processed_scope = get_target_elements(doc, category_bic)
    if not elements:
        ui.uiUtils_alert("No {}s were found to map.".format(category_label.lower()), title=TITLE)
        return

    mapped_count = 0
    already_count = 0
    missing_param_count = 0
    unmatched_names = []

    with revit.Transaction(TITLE):
        for elem in elements:
            host_name = get_target_name(elem, category_bic)
            host_name_norm = _normalize_name(host_name)
            if not host_name_norm:
                unmatched_names.append("(blank name)")
                continue

            match = chosen_by_name.get(host_name_norm)
            if match is None:
                unmatched_names.append(host_name)
                continue

            try:
                param = elem.LookupParameter(host_key_param_name)
            except Exception:
                param = None
            if not param:
                missing_param_count += 1
                continue
            try:
                if param.IsReadOnly or param.StorageType != DB.StorageType.ElementId:
                    missing_param_count += 1
                    continue
            except Exception:
                missing_param_count += 1
                continue

            target_id = getattr(match.get("element"), "Id", None)
            if target_id is None:
                continue
            try:
                current_id = param.AsElementId()
            except Exception:
                current_id = None
            if _element_id_value(current_id) == _element_id_value(target_id):
                already_count += 1
                continue
            try:
                param.Set(target_id)
                mapped_count += 1
            except Exception:
                missing_param_count += 1

    unique_unmatched = []
    seen_unmatched = set()
    for name in unmatched_names:
        normalized = _normalize_name(name)
        if normalized in seen_unmatched:
            continue
        seen_unmatched.add(normalized)
        unique_unmatched.append(name)

    _safe_config_set("map_key_schedule_target", selected_target_type)
    _safe_config_set("map_key_schedule_name_{}".format(category_key), schedule.Name)
    _save_config()

    ui.uiUtils_alert(
        build_summary(
            processed_scope,
            category_label,
            schedule.Name,
            mapped_count,
            already_count,
            unique_unmatched,
            missing_param_count,
            duplicate_samples,
            resolved_to_residential,
        ),
        title=TITLE,
    )


if __name__ == "__main__":
    main()
