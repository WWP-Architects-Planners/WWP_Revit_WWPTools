#!python3
# -*- coding: utf-8 -*-

import os
import sys
import importlib
from pyrevit import revit, DB
import WWP_uiUtils as ui

script_dir = os.path.dirname(__file__)
pulldown_dir = os.path.abspath(os.path.join(script_dir, ".."))
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "..", "lib"))
for path in (pulldown_dir, lib_path):
    if path not in sys.path:
        sys.path.append(path)

for module_name in ("WWP_colorSchemeUtils", "color_scheme_common"):
    try:
        if module_name in sys.modules:
            del sys.modules[module_name]
    except Exception:
        pass

import WWP_colorSchemeUtils as csu
import color_scheme_common as csc

try:
    csu = importlib.reload(csu)
except Exception:
    pass

try:
    csc = importlib.reload(csc)
except Exception:
    pass


def _log(message):
    try:
        print("[Import Color Scheme] {}".format(message))
    except Exception:
        pass


def _elem_id_int(elem_id):
    try:
        return int(elem_id.IntegerValue)
    except Exception:
        try:
            return int(elem_id)
        except Exception:
            return None


def _category_name(doc, category_id):
    if category_id is None:
        return "Unknown Category"
    cat_int = _elem_id_int(category_id)
    try:
        categories = getattr(getattr(doc, "Settings", None), "Categories", None)
        if categories is not None:
            cat = categories.get_Item(category_id)
            cat_name = getattr(cat, "Name", "") if cat else ""
            if cat_name:
                return cat_name
    except Exception:
        pass
    try:
        import System
        bic = System.Enum.ToObject(DB.BuiltInCategory, cat_int)
        categories = getattr(getattr(doc, "Settings", None), "Categories", None)
        if categories is not None:
            cat = categories.get_Item(bic)
            cat_name = getattr(cat, "Name", "") if cat else ""
            if cat_name:
                return cat_name
        label = DB.LabelUtils.GetLabelFor(bic)
        if label:
            return label
    except Exception:
        pass
    try:
        label = DB.LabelUtils.GetLabelFor(category_id)
        if label:
            return label
    except Exception:
        pass
    if cat_int is not None:
        return "Category {}".format(cat_int)
    return "Unknown Category"


def _scope_label(doc, scheme):
    area_name = csc.scheme_area_scheme_name(doc, scheme)
    if area_name:
        return "Area({})".format(area_name)
    category_label = _category_name(doc, getattr(scheme, "CategoryId", None))
    if not category_label or category_label == "Unknown Category":
        return ""
    return category_label


def _build_choices(doc, schemes):
    scope_map = {}
    for scheme in schemes:
        key = (
            _elem_id_int(getattr(scheme, "CategoryId", None)),
            _elem_id_int(csc.scheme_area_scheme_id(scheme)),
        )
        if key not in scope_map:
            scope_map[key] = scheme
        if _scope_label(doc, scheme) == "":
            _log(
                "Unresolved category for scheme '{}' categoryId={} areaSchemeId={}".format(
                    getattr(scheme, "Name", "") or "Color Scheme",
                    _elem_id_int(getattr(scheme, "CategoryId", None)),
                    _elem_id_int(csc.scheme_area_scheme_id(scheme)),
                )
            )
    create_choices = [{
        "mode": "create",
        "label": "Create New In {}".format(_scope_label(doc, scheme) or (getattr(scheme, "Name", "") or "Selected Scope")),
        "scheme": scheme,
    } for scheme in scope_map.values()]
    overwrite_choices = [{
        "mode": "overwrite",
        "label": (
            "Overwrite {}: {}".format(_scope_label(doc, scheme), getattr(scheme, "Name", "") or "Color Scheme")
            if _scope_label(doc, scheme)
            else "Overwrite {}".format(getattr(scheme, "Name", "") or "Color Scheme")
        ),
        "scheme": scheme,
    } for scheme in schemes]
    create_choices.sort(key=lambda x: x["label"].lower())
    overwrite_choices.sort(key=lambda x: x["label"].lower())
    return create_choices + overwrite_choices


def _format_payload_snapshot(payload):
    storage_types = sorted(set([
        str(item.get("storage_type", "")).strip()
        for item in payload.get("entries", [])
        if str(item.get("storage_type", "")).strip()
    ]))
    lines = [
        "Workbook",
        "Scheme: {}".format(payload.get("scheme_name", "")),
        "Category: {}".format(payload.get("category_name", "")),
        "Area Scheme: {}".format(payload.get("area_scheme_name", "")),
        "Title: {}".format(payload.get("title", "")),
        "Parameter: {} ({})".format(payload.get("parameter_name", ""), payload.get("parameter_id", "")),
        "Modes: ByValue={} ByRange={} ByPercentage={}".format(
            bool(payload.get("is_by_value")),
            bool(payload.get("is_by_range")),
            bool(payload.get("is_by_percentage")),
        ),
        "Entry Storage Types: {}".format(", ".join(storage_types) if storage_types else "<none>"),
        "Entry Count: {}".format(len(payload.get("entries", []))),
    ]
    return "\n".join(lines)


def _format_target_snapshot(snapshot):
    lines = [
        "Target Scheme",
        "Scheme: {}".format(snapshot.get("scheme_name", "")),
        "Category: {}".format(snapshot.get("category_name", "")),
        "Area Scheme: {}".format(snapshot.get("area_scheme_name", "")),
        "Title: {}".format(snapshot.get("title", "")),
        "Parameter: {} ({})".format(snapshot.get("parameter_name", ""), snapshot.get("parameter_id", "")),
        "Modes: ByValue={} ByRange={} ByPercentage={}".format(
            bool(snapshot.get("is_by_value")),
            bool(snapshot.get("is_by_range")),
            bool(snapshot.get("is_by_percentage")),
        ),
        "Entry Storage Types: {}".format(", ".join(snapshot.get("entry_storage_types", [])) if snapshot.get("entry_storage_types") else "<none>"),
    ]
    return "\n".join(lines)


def main():
    path = ui.uiUtils_open_file_dialog(
        title="Import Color Scheme from Excel",
        filter_text="Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All Files (*.*)|*.*",
        multiselect=False,
    )
    if not path:
        return

    try:
        payload = csc.import_payload_from_excel(path)
    except Exception as ex:
        ui.uiUtils_alert("Failed to read color scheme workbook.\n\n{}".format(str(ex)), title="Import Color Scheme")
        return

    workbook_summary = _format_payload_snapshot(payload)
    if not ui.uiUtils_show_text_report(
        "Import Color Scheme",
        workbook_summary,
        ok_text="Continue",
        cancel_text="Cancel",
        width=760,
        height=420,
    ):
        return

    doc = revit.doc
    schemes = csu.collect_color_fill_schemes(doc)
    if not schemes:
        ui.uiUtils_alert("No Color Fill Schemes found in this model.", title="Import Color Scheme")
        return

    choices = _build_choices(doc, schemes)
    if not choices:
        ui.uiUtils_alert("No target categories or schemes are available for import.", title="Import Color Scheme")
        return

    labels = [choice["label"] for choice in choices]
    prechecked = [csc.choose_default_import_target_index(doc, payload, choices)]
    selected = ui.uiUtils_select_indices(
        labels,
        title="Import Color Scheme",
        prompt="Select the target category/scope or an existing scheme to overwrite:",
        multiselect=False,
        width=860,
        height=620,
    )
    if not selected:
        return

    choice = choices[int(selected[0])]
    target = choice["scheme"]
    imported_name = payload.get("scheme_name", "Imported Color Scheme")
    ok_preflight, problems, target_snapshot = csc.validate_payload_against_scheme(doc, payload, target)
    _log("Preflight target scheme: {}".format(target_snapshot))
    comparison_text = _format_payload_snapshot(payload) + "\n\n" + _format_target_snapshot(target_snapshot)
    if problems:
        comparison_text += "\n\nMismatch\n" + "\n".join(problems)
        ui.uiUtils_show_text_report(
            "Import Color Scheme",
            comparison_text,
            ok_text="Close",
            cancel_text=None,
            width=860,
            height=560,
        )
        return
    if not ui.uiUtils_show_text_report(
        "Import Color Scheme",
        comparison_text,
        ok_text="Import",
        cancel_text="Cancel",
        width=860,
        height=560,
    ):
        return

    with revit.Transaction("Import Color Scheme From Excel"):
        if choice["mode"] == "create":
            new_name = csc.unique_scheme_name_in_scope(doc, target, imported_name)
            new_id = target.Duplicate(new_name)
            target = doc.GetElement(new_id)
            if target is None:
                ui.uiUtils_alert("Failed to create a new color scheme in the selected target scope.", title="Import Color Scheme")
                return
        ok, error = csc.apply_payload_to_scheme(target, payload, log=_log)

    if not ok:
        ui.uiUtils_alert("Failed to import color scheme.\n\n{}".format(error or "Unknown error"), title="Import Color Scheme")
        return

    try:
        revit.get_selection().set_to([target.Id])
    except Exception:
        pass

    ui.uiUtils_alert("Imported scheme to '{}'.".format(getattr(target, "Name", "Color Scheme")), title="Import Color Scheme")


if __name__ == "__main__":
    main()
