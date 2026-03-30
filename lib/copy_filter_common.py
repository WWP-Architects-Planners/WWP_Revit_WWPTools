import ast
import os

import clr

from Autodesk.Revit import DB

_WPFUI_THEME_READY = False


def normalize_text(value):
    return " ".join(str(value or "").split()).strip()


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


def to_element_id(value):
    return DB.ElementId(int(value))


def get_active_view(uidoc):
    try:
        return uidoc.ActiveView if uidoc else None
    except Exception:
        return None


def is_filterable_view(view):
    if view is None:
        return False
    try:
        if view.ViewType == DB.ViewType.Schedule:
            return False
    except Exception:
        pass
    try:
        list(get_ordered_filter_ids(view))
        return True
    except Exception:
        return False


def get_ordered_filter_ids(view):
    getter = getattr(view, "GetOrderedFilters", None)
    if callable(getter):
        try:
            return list(getter())
        except Exception:
            pass
    return list(view.GetFilters())


def view_display_name(view):
    view_type = normalize_text(getattr(view, "ViewType", "View"))
    name = normalize_text(getattr(view, "Name", "")) or "(unnamed)"
    if getattr(view, "IsTemplate", False):
        return "{} | {}".format(view_type, name)
    return "{} | {}".format(view_type, name)


def filter_display_name(filter_element, is_visible):
    name = normalize_text(getattr(filter_element, "Name", "")) or "(unnamed filter)"
    if is_visible:
        return name
    return "{}  [Hidden]".format(name)


def collect_template_views(doc):
    templates = []
    collector = DB.FilteredElementCollector(doc).OfClass(DB.View)
    for view in collector:
        try:
            if not view.IsTemplate:
                continue
        except Exception:
            continue
        if not is_filterable_view(view):
            continue
        templates.append(view)
    templates.sort(key=lambda item: (normalize_text(getattr(item, "ViewType", "")).lower(), normalize_text(getattr(item, "Name", "")).lower()))
    return templates


def collect_filter_records(doc, view):
    records = []
    if not is_filterable_view(view):
        return records
    for index_value, filter_id in enumerate(get_ordered_filter_ids(view), 1):
        filter_element = doc.GetElement(filter_id)
        if filter_element is None:
            continue
        try:
            is_visible = bool(view.GetFilterVisibility(filter_id))
        except Exception:
            is_visible = True
        records.append({
            "id": element_id_value(filter_id),
            "filter_id": filter_id,
            "filter_element": filter_element,
            "name": normalize_text(getattr(filter_element, "Name", "")),
            "display": filter_display_name(filter_element, is_visible),
            "is_visible": is_visible,
            "order": index_value,
        })
    return records


def build_filter_record_lookup(records):
    return {record["id"]: record for record in (records or []) if record.get("id") is not None}


def source_view_record(view):
    return {
        "id": element_id_value(getattr(view, "Id", None)),
        "name": normalize_text(getattr(view, "Name", "")),
        "display": view_display_name(view),
        "view": view,
    }


def template_view_records(doc, exclude_view_id=None):
    excluded = element_id_value(exclude_view_id)
    records = []
    for view in collect_template_views(doc):
        view_id = element_id_value(view.Id)
        if excluded is not None and view_id == excluded:
            continue
        records.append({
            "id": view_id,
            "name": normalize_text(getattr(view, "Name", "")),
            "display": view_display_name(view),
            "view": view,
        })
    return records


def _copy_override_graphics(source_view, target_view, filter_id):
    overrides = source_view.GetFilterOverrides(filter_id)
    target_view.SetFilterOverrides(filter_id, overrides)
    try:
        target_view.SetFilterVisibility(filter_id, bool(source_view.GetFilterVisibility(filter_id)))
    except Exception:
        pass


def copy_filters_to_targets(doc, source_view, target_views, filter_ids, transaction_name, clear_target_view_template=False):
    selected_ids = [int(filter_id) for filter_id in (filter_ids or [])]
    if not selected_ids:
        raise Exception("Select at least one filter.")

    filter_lookup = build_filter_record_lookup(collect_filter_records(doc, source_view))
    missing_ids = [filter_id for filter_id in selected_ids if filter_id not in filter_lookup]
    if missing_ids:
        raise Exception("One or more selected filters are no longer available in the source view.")

    transaction = DB.Transaction(doc, transaction_name)
    transaction.Start()
    updated = []
    try:
        for target_view in target_views:
            if target_view is None:
                continue
            if clear_target_view_template and not getattr(target_view, "IsTemplate", False):
                try:
                    if target_view.ViewTemplateId != DB.ElementId.InvalidElementId:
                        target_view.ViewTemplateId = DB.ElementId.InvalidElementId
                except Exception:
                    pass

            existing_ids = set(element_id_value(filter_id) for filter_id in get_ordered_filter_ids(target_view))
            for selected_id in selected_ids:
                filter_id = to_element_id(selected_id)
                if selected_id not in existing_ids:
                    target_view.AddFilter(filter_id)
                    existing_ids.add(selected_id)
                _copy_override_graphics(source_view, target_view, filter_id)
            updated.append(target_view)
        transaction.Commit()
    except Exception:
        if transaction.HasStarted():
            transaction.RollBack()
        raise
    return updated


def read_bundle_title(script_dir, default_title):
    bundle_path = os.path.join(script_dir, "bundle.yaml")
    if not os.path.isfile(bundle_path):
        return default_title
    try:
        with open(bundle_path, "r") as bundle_file:
            for raw_line in bundle_file:
                line = raw_line.strip()
                if not line.lower().startswith("title:"):
                    continue
                value = line.split(":", 1)[1].strip()
                if not value:
                    break
                try:
                    parsed = ast.literal_eval(value)
                    if parsed:
                        return str(parsed)
                except Exception:
                    return value.strip("\"'")
    except Exception:
        pass
    return default_title


def ensure_wpfui_theme(lib_path):
    global _WPFUI_THEME_READY
    if _WPFUI_THEME_READY:
        return
    try:
        revit_version = int(str(__revit__.Application.VersionNumber))
    except Exception:
        revit_version = None
    dll_name = "WWPTools.WpfUI.net8.0-windows.dll" if revit_version and revit_version >= 2025 else "WWPTools.WpfUI.net48.dll"
    dll_path = os.path.join(lib_path, dll_name)
    if not os.path.isfile(dll_path):
        return
    try:
        if hasattr(clr, "AddReferenceToFileAndPath"):
            clr.AddReferenceToFileAndPath(dll_path)
        else:
            clr.AddReference(dll_path)
        _WPFUI_THEME_READY = True
    except Exception:
        pass
