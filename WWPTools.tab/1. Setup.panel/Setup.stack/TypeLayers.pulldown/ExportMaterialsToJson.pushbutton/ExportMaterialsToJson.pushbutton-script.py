import clr
import json
import os
import sys

from pyrevit import DB, HOST_APP, script


def get_doc():
    doc = HOST_APP.doc
    if doc is None:
        raise Exception("No active Revit document.")
    return doc


def add_lib_path():
    script_dir = os.path.dirname(__file__)
    lib_path = None
    current = os.path.abspath(script_dir)
    for _ in range(8):
        candidate = os.path.join(current, "lib")
        if os.path.isdir(candidate):
            lib_path = candidate
            break
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    if lib_path is None:
        lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "..", "lib"))
    if lib_path not in sys.path:
        sys.path.append(lib_path)


def load_uiutils():
    add_lib_path()
    import WWP_uiUtils as ui
    return ui


def get_elem_name(elem):
    # Some Revit API versions expose Element.Name in a way IronPython's dynamic
    # attribute lookup cannot resolve directly for certain element types
    # (fails with "AttributeError: Name" / MissingMemberException even though
    # the member exists). Fall back to the reflected property descriptor and,
    # failing that, the underlying type-name parameter.
    if elem is None:
        return None
    try:
        return elem.Name
    except Exception:
        pass
    try:
        return DB.Element.Name.__get__(elem)
    except Exception:
        pass
    try:
        bips = (DB.BuiltInParameter.SYMBOL_NAME_PARAM, DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
    except Exception:
        bips = ()
    for bip in bips:
        try:
            param = elem.get_Parameter(bip)
            if param and param.HasValue:
                value = param.AsString()
                if value:
                    return value
        except Exception:
            continue
    return None


def color_to_list(color):
    try:
        if color is None or not color.IsValid:
            return None
        # Color.Red/Green/Blue come back as raw .NET Byte values, which
        # Python's json module does not recognize as serializable ints.
        return [int(color.Red), int(color.Green), int(color.Blue)]
    except Exception:
        return None


def pattern_name(doc, pattern_id):
    if pattern_id is None or pattern_id == DB.ElementId.InvalidElementId:
        return None
    try:
        pattern_elem = doc.GetElement(pattern_id)
    except Exception:
        pattern_elem = None
    if pattern_elem is None:
        return None
    return get_elem_name(pattern_elem)


def get_material_keynote(mat):
    try:
        param = mat.get_Parameter(DB.BuiltInParameter.KEYNOTE_PARAM)
        if param and param.HasValue:
            return param.AsString()
    except Exception:
        pass
    return None


def material_to_dict(doc, mat):
    name = get_elem_name(mat)
    if not name:
        return None

    record = {"name": name}

    try:
        record["materialClass"] = mat.MaterialClass
    except Exception:
        record["materialClass"] = None
    try:
        record["materialCategory"] = mat.MaterialCategory
    except Exception:
        record["materialCategory"] = None
    try:
        record["color"] = color_to_list(mat.Color)
    except Exception:
        record["color"] = None
    try:
        record["transparency"] = int(mat.Transparency)
    except Exception:
        record["transparency"] = None
    try:
        record["shininess"] = int(mat.Shininess)
    except Exception:
        record["shininess"] = None
    try:
        record["smoothness"] = int(mat.Smoothness)
    except Exception:
        record["smoothness"] = None

    try:
        record["cutForegroundPattern"] = pattern_name(doc, mat.CutForegroundPatternId)
    except Exception:
        record["cutForegroundPattern"] = None
    try:
        record["cutForegroundColor"] = color_to_list(mat.CutForegroundPatternColor)
    except Exception:
        record["cutForegroundColor"] = None
    try:
        record["cutBackgroundPattern"] = pattern_name(doc, mat.CutBackgroundPatternId)
    except Exception:
        record["cutBackgroundPattern"] = None
    try:
        record["cutBackgroundColor"] = color_to_list(mat.CutBackgroundPatternColor)
    except Exception:
        record["cutBackgroundColor"] = None

    try:
        record["surfaceForegroundPattern"] = pattern_name(doc, mat.SurfaceForegroundPatternId)
    except Exception:
        record["surfaceForegroundPattern"] = None
    try:
        record["surfaceForegroundColor"] = color_to_list(mat.SurfaceForegroundPatternColor)
    except Exception:
        record["surfaceForegroundColor"] = None
    try:
        record["surfaceBackgroundPattern"] = pattern_name(doc, mat.SurfaceBackgroundPatternId)
    except Exception:
        record["surfaceBackgroundPattern"] = None
    try:
        record["surfaceBackgroundColor"] = color_to_list(mat.SurfaceBackgroundPatternColor)
    except Exception:
        record["surfaceBackgroundColor"] = None

    record["keynote"] = get_material_keynote(mat)
    return record


def pick_export_path(doc, ui):
    initial_dir = ""
    try:
        if doc.PathName:
            initial_dir = os.path.dirname(doc.PathName)
    except Exception:
        initial_dir = ""
    if not initial_dir or not os.path.isdir(initial_dir):
        initial_dir = ""
    path = ui.uiUtils_save_file_dialog(
        title="Export Materials To JSON",
        filter_text="JSON File (*.json)|*.json",
        default_extension="json",
        initial_directory=initial_dir,
        file_name="Materials.json",
    )
    if not path:
        return None
    if not path.lower().endswith(".json"):
        path = "{}.json".format(path)
    return path


def choose_materials_to_export(ui, records):
    items = [record.get("name") or "" for record in records]
    selected_indices, _filter_param, _filter_value = ui.uiUtils_select_items_with_filter(
        items,
        [],
        [],
        title="Export Materials To JSON",
        prompt="Select materials to export. Type in Search to filter live; Select All / None / Invert to adjust.",
        prechecked_indices=list(range(len(items))),
        width=900,
        height=700,
    )
    if not selected_indices:
        return []
    return [records[idx] for idx in selected_indices]


def main():
    doc = get_doc()
    ui = load_uiutils()

    materials = list(DB.FilteredElementCollector(doc).OfClass(DB.Material))
    records = []
    for mat in materials:
        record = material_to_dict(doc, mat)
        if record:
            records.append(record)
    records.sort(key=lambda r: (r.get("name") or "").lower())

    if not records:
        ui.uiUtils_alert("No materials found in this project.", title="Export Materials To JSON")
        return

    selected_records = choose_materials_to_export(ui, records)
    if not selected_records:
        return

    file_path = pick_export_path(doc, ui)
    if not file_path:
        return

    payload = {
        "exportedFrom": doc.PathName or doc.Title,
        "materialCount": len(selected_records),
        "materials": selected_records,
    }

    try:
        with open(file_path, "w") as f:
            json.dump(payload, f, indent=2)
    except Exception as exc:
        ui.uiUtils_alert("Failed to write JSON file.\n{}".format(exc), title="Export Materials To JSON")
        return

    ui.uiUtils_alert(
        "Exported {} material(s) to:\n{}".format(len(selected_records), file_path),
        title="Export Materials To JSON",
    )


if __name__ == "__main__":
    main()
