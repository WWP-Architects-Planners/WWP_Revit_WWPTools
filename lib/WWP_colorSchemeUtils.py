#!python3
# -*- coding: utf-8 -*-

from pyrevit import DB

try:
    from System.Collections.Generic import List
except Exception:
    List = None


def _elem_id_int(elem_id):
    try:
        return int(elem_id.IntegerValue)
    except Exception:
        try:
            return int(elem_id)
        except Exception:
            return None


def _scheme_area_scheme_id(scheme):
    try:
        area_scheme_id = getattr(scheme, "AreaSchemeId", None)
        if area_scheme_id and area_scheme_id != DB.ElementId.InvalidElementId:
            return area_scheme_id
    except Exception:
        pass
    try:
        getter = getattr(scheme, "GetAreaSchemeId", None)
        if callable(getter):
            area_scheme_id = getter()
            if area_scheme_id and area_scheme_id != DB.ElementId.InvalidElementId:
                return area_scheme_id
    except Exception:
        pass
    return None


def _same_scope(a, b):
    if a is None or b is None:
        return False
    return (
        _elem_id_int(getattr(a, "CategoryId", None)) == _elem_id_int(getattr(b, "CategoryId", None))
        and _elem_id_int(_scheme_area_scheme_id(a)) == _elem_id_int(_scheme_area_scheme_id(b))
    )


def _collect_color_fill_schemes(doc):
    try:
        schemes = list(DB.FilteredElementCollector(doc).OfClass(DB.ColorFillScheme).ToElements())
    except Exception:
        schemes = []
    return [s for s in schemes if s is not None]


def collect_color_fill_schemes(doc):
    return _collect_color_fill_schemes(doc)


def _find_scheme_in_scope_by_name(all_schemes, scope_seed_scheme, name):
    target_name = (name or "").strip().lower()
    if not target_name:
        return None
    for s in all_schemes:
        try:
            if not _same_scope(s, scope_seed_scheme):
                continue
            if (getattr(s, "Name", "") or "").strip().lower() == target_name:
                return s
        except Exception:
            continue
    return None


def find_scheme_in_scope_by_name(all_schemes, scope_seed_scheme, name):
    return _find_scheme_in_scope_by_name(all_schemes, scope_seed_scheme, name)


def _entry_get_value(entry, storage_type):
    try:
        if storage_type == DB.StorageType.String:
            for name in ("GetStringValue", "AsString"):
                fn = getattr(entry, name, None)
                if callable(fn):
                    return fn()
            return getattr(entry, "StringValue", None)
        if storage_type == DB.StorageType.Integer:
            for name in ("GetIntegerValue", "AsInteger"):
                fn = getattr(entry, name, None)
                if callable(fn):
                    return fn()
            return getattr(entry, "IntegerValue", None)
        if storage_type == DB.StorageType.Double:
            for name in ("GetDoubleValue", "AsDouble"):
                fn = getattr(entry, name, None)
                if callable(fn):
                    return fn()
            return getattr(entry, "DoubleValue", None)
        if storage_type == DB.StorageType.ElementId:
            for name in ("GetElementIdValue", "AsElementId"):
                fn = getattr(entry, name, None)
                if callable(fn):
                    return fn()
            return getattr(entry, "ElementIdValue", None)
    except Exception:
        return None
    return None


def _entry_set_value(entry, storage_type, value):
    try:
        if storage_type == DB.StorageType.String:
            for name in ("SetStringValue",):
                fn = getattr(entry, name, None)
                if callable(fn):
                    fn("" if value is None else str(value))
                    return True
            if hasattr(entry, "StringValue"):
                entry.StringValue = "" if value is None else str(value)
                return True
        if storage_type == DB.StorageType.Integer:
            for name in ("SetIntegerValue",):
                fn = getattr(entry, name, None)
                if callable(fn):
                    fn(int(value if value is not None else 0))
                    return True
            if hasattr(entry, "IntegerValue"):
                entry.IntegerValue = int(value if value is not None else 0)
                return True
        if storage_type == DB.StorageType.Double:
            for name in ("SetDoubleValue",):
                fn = getattr(entry, name, None)
                if callable(fn):
                    fn(float(value if value is not None else 0.0))
                    return True
            if hasattr(entry, "DoubleValue"):
                entry.DoubleValue = float(value if value is not None else 0.0)
                return True
        if storage_type == DB.StorageType.ElementId:
            if not isinstance(value, DB.ElementId):
                value = DB.ElementId.InvalidElementId
            for name in ("SetElementIdValue",):
                fn = getattr(entry, name, None)
                if callable(fn):
                    fn(value)
                    return True
            if hasattr(entry, "ElementIdValue"):
                entry.ElementIdValue = value
                return True
    except Exception:
        return False
    return False


def _entry_caption(entry):
    try:
        getter = getattr(entry, "GetCaption", None)
        if callable(getter):
            val = getter()
            if val:
                return str(val)
    except Exception:
        pass
    try:
        val = getattr(entry, "Caption", None)
        if val:
            return str(val)
    except Exception:
        pass
    return ""


def _build_custom_entry_from_source(source_entry):
    try:
        storage_type = getattr(source_entry, "StorageType", None)
        if storage_type is None:
            return None
        custom = DB.ColorFillSchemeEntry(storage_type)
    except Exception:
        return None

    value = _entry_get_value(source_entry, storage_type)
    if storage_type == DB.StorageType.ElementId:
        value = DB.ElementId.InvalidElementId
    if not _entry_set_value(custom, storage_type, value):
        return None

    for attr in ("Color", "FillPatternId", "IsVisible", "Visible"):
        if hasattr(source_entry, attr) and hasattr(custom, attr):
            try:
                setattr(custom, attr, getattr(source_entry, attr))
            except Exception:
                pass

    caption = _entry_caption(source_entry)
    if caption:
        try:
            setter = getattr(custom, "SetCaption", None)
            if callable(setter):
                setter(caption)
            elif hasattr(custom, "Caption"):
                custom.Caption = caption
        except Exception:
            pass
    return custom


def _same_entry_value(a, b):
    try:
        st_a = getattr(a, "StorageType", None)
        st_b = getattr(b, "StorageType", None)
        if st_a != st_b:
            return False
        va = _entry_get_value(a, st_a)
        vb = _entry_get_value(b, st_b)
        if st_a == DB.StorageType.ElementId:
            try:
                ai = int(va.IntegerValue)
                bi = int(vb.IntegerValue)
                return ai == bi
            except Exception:
                return False
        return va == vb
    except Exception:
        return False


def _overwrite_existing_matching_entry(target, source_entry):
    try:
        entries = list(target.GetEntries())
    except Exception:
        entries = []
    src_caption = _entry_caption(source_entry).strip().lower()
    for existing in entries:
        try:
            same = _same_entry_value(existing, source_entry)
            if not same and src_caption:
                ex_caption = _entry_caption(existing).strip().lower()
                same = ex_caption and ex_caption == src_caption
            if not same:
                continue
            for attr in ("Color", "FillPatternId", "IsVisible", "Visible"):
                if hasattr(source_entry, attr) and hasattr(existing, attr):
                    try:
                        setattr(existing, attr, getattr(source_entry, attr))
                    except Exception:
                        pass
            return True
        except Exception:
            continue
    return False


def copy_scheme_data(source, target):
    for attr in ("Name", "Title", "IsByRange", "IsByValue", "IsByPercentage"):
        if hasattr(source, attr) and hasattr(target, attr):
            try:
                setattr(target, attr, getattr(source, attr))
            except Exception:
                pass
    for attr in ("ParameterDefinition", "ParameterId"):
        if hasattr(source, attr) and hasattr(target, attr):
            try:
                setattr(target, attr, getattr(source, attr))
            except Exception:
                pass

    try:
        source_entries = list(source.GetEntries())
    except Exception:
        return False, "Unable to read entries from source scheme."

    try:
        clear_entries = getattr(target, "ClearEntries", None)
        if callable(clear_entries):
            clear_entries()
        else:
            remove_entry = getattr(target, "RemoveEntry", None)
            if callable(remove_entry):
                for entry in list(target.GetEntries()):
                    remove_entry(entry)
    except Exception:
        pass

    set_entries_applied = False
    try:
        set_entries = getattr(target, "SetEntries", None)
        if callable(set_entries):
            if List is not None:
                set_entries(List[DB.ColorFillSchemeEntry](source_entries))
            else:
                set_entries(list(source_entries))
            try:
                if len(list(target.GetEntries())) == len(source_entries):
                    set_entries_applied = True
                    return True, None
            except Exception:
                set_entries_applied = True
                return True, None
    except Exception:
        pass

    if not set_entries_applied:
        try:
            clear_entries = getattr(target, "ClearEntries", None)
            if callable(clear_entries):
                clear_entries()
            else:
                remove_entry = getattr(target, "RemoveEntry", None)
                if callable(remove_entry):
                    for entry in list(target.GetEntries()):
                        remove_entry(entry)
        except Exception:
            pass

    add_entry = getattr(target, "AddEntry", None)
    if not callable(add_entry):
        return False, "Target scheme API does not support SetEntries/AddEntry."

    added = 0
    failures = []
    for entry in source_entries:
        try:
            clone = getattr(entry, "Clone", None)
            new_entry = clone() if callable(clone) else entry
            add_entry(new_entry)
            added += 1
        except Exception as ex_add:
            try:
                if _overwrite_existing_matching_entry(target, entry):
                    added += 1
                    continue
            except Exception:
                pass
            try:
                custom_entry = _build_custom_entry_from_source(entry)
                if custom_entry is not None:
                    add_entry(custom_entry)
                    added += 1
                    continue
            except Exception as ex:
                if len(failures) < 5:
                    failures.append(str(ex))
            if len(failures) < 5:
                failures.append(str(ex_add) or "failed to add entry")

    if added < len(source_entries):
        detail = (" Errors: " + " | ".join(failures)) if failures else ""
        return False, "Only copied {} of {} entries.{}".format(added, len(source_entries), detail)
    return True, None


def _get_view_color_fill_scheme_id(view, category_id):
    try:
        method = getattr(view, "GetColorFillSchemeId", None)
        if callable(method):
            return method(category_id)
    except Exception:
        pass
    return None


def _set_view_color_fill_scheme_id(view, category_id, scheme_id):
    try:
        method = getattr(view, "SetColorFillSchemeId", None)
        if callable(method):
            method(category_id, scheme_id)
            return True
    except Exception:
        pass
    return False


def _list_target_area_color_schemes(doc, target_area_scheme_id):
    area_category_id = DB.ElementId(DB.BuiltInCategory.OST_Areas)
    matches = []
    for scheme in _collect_color_fill_schemes(doc):
        try:
            if _elem_id_int(_scheme_area_scheme_id(scheme)) != _elem_id_int(target_area_scheme_id):
                continue
            if _elem_id_int(getattr(scheme, "CategoryId", None)) != _elem_id_int(area_category_id):
                continue
            matches.append(scheme)
        except Exception:
            continue
    return matches


def copy_view_area_color_scheme_with_scope(doc, source_view, target_view, target_area_scheme_id):
    area_category_id = DB.ElementId(DB.BuiltInCategory.OST_Areas)
    source_scheme_id = _get_view_color_fill_scheme_id(source_view, area_category_id)
    if source_scheme_id is None or source_scheme_id == DB.ElementId.InvalidElementId:
        return False, "source view has no area color scheme assignment"
    source_scheme = doc.GetElement(source_scheme_id)
    if source_scheme is None:
        return False, "source color scheme element not found"

    # Scope seed: current target view scheme if valid in target area scheme, else first scheme in that scope.
    scope_seed = None
    target_current_id = _get_view_color_fill_scheme_id(target_view, area_category_id)
    if target_current_id is not None and target_current_id != DB.ElementId.InvalidElementId:
        try:
            current_target = doc.GetElement(target_current_id)
            if current_target is not None and _elem_id_int(_scheme_area_scheme_id(current_target)) == _elem_id_int(target_area_scheme_id):
                scope_seed = current_target
        except Exception:
            pass
    if scope_seed is None:
        candidates = _list_target_area_color_schemes(doc, target_area_scheme_id)
        if not candidates:
            return False, "target area scheme has no area color scheme to duplicate from"
        scope_seed = candidates[0]

    source_name = (getattr(source_scheme, "Name", "") or "Color Scheme").strip()
    selected_name = (getattr(scope_seed, "Name", "") or "").strip()
    target_scheme = scope_seed
    if selected_name.lower() != source_name.lower():
        all_schemes = _collect_color_fill_schemes(doc)
        existing_same_name = _find_scheme_in_scope_by_name(all_schemes, scope_seed, source_name)
        if existing_same_name is not None:
            target_scheme = existing_same_name
        else:
            try:
                new_id = scope_seed.Duplicate(source_name)
                created = doc.GetElement(new_id)
            except Exception as ex:
                created = None
                err = str(ex)
            if created is None:
                return False, "failed to create new scheme '{}' in target scope{}".format(
                    source_name,
                    (" ({})".format(err) if 'err' in locals() and err else ""),
                )
            target_scheme = created

    ok, msg = copy_scheme_data(source_scheme, target_scheme)
    if not ok:
        return False, msg
    if not _set_view_color_fill_scheme_id(target_view, area_category_id, target_scheme.Id):
        return False, "failed to assign target color scheme to target view ({})".format(
            getattr(target_scheme, "Name", "<unknown>")
        )
    return True, ""
