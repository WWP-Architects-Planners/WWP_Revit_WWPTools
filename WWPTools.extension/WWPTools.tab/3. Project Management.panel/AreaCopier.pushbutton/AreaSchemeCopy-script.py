#! python3
import os
import sys
import traceback

from pyrevit import DB, revit

script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.append(lib_path)

def _load_uiutils():
    try:
        import WWP_uiUtils as ui
        return ui
    except Exception:
        try:
            from pyrevit import forms
            forms.alert(
                "WWP_uiUtils is not available. Restart pyRevit or reinstall WWPTools.",
                title="Copy Areas",
            )
        except Exception:
            pass
        raise


def _get_area_schemes(doc):
    schemes = list(DB.FilteredElementCollector(doc).OfClass(DB.AreaScheme))
    schemes = [s for s in schemes if s is not None]
    schemes.sort(key=lambda s: (s.Name or "").lower())
    return schemes


def _get_levels(doc):
    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level).WhereElementIsNotElementType())
    levels.sort(key=lambda l: l.Elevation)
    return levels


def _get_area_plans_by_level(doc, scheme_id):
    plans = {}
    views = DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan)
    for view in views:
        try:
            if view.ViewType != DB.ViewType.AreaPlan:
                continue
        except Exception:
            continue
        if view.IsTemplate:
            continue

        view_scheme_id = None
        try:
            if hasattr(view, "AreaSchemeId"):
                view_scheme_id = view.AreaSchemeId
        except Exception:
            view_scheme_id = None
        if view_scheme_id is None:
            try:
                if hasattr(view, "AreaScheme") and view.AreaScheme:
                    view_scheme_id = view.AreaScheme.Id
            except Exception:
                view_scheme_id = None
        if view_scheme_id is None:
            try:
                param = view.get_Parameter(DB.BuiltInParameter.VIEW_AREA_SCHEME)
                if param:
                    view_scheme_id = param.AsElementId()
            except Exception:
                view_scheme_id = None

        if view_scheme_id is None or view_scheme_id != scheme_id:
            continue

        level_id = None
        try:
            if view.GenLevel:
                level_id = view.GenLevel.Id
        except Exception:
            level_id = None
        if level_id is None:
            try:
                level_id = view.LevelId
            except Exception:
                level_id = None
        if level_id is None:
            try:
                level_id = view.AssociatedLevelId
            except Exception:
                level_id = None

        if level_id is not None and level_id not in plans:
            plans[level_id] = view
    return plans


def _get_area_boundary_category_ids():
    ids = set()
    for name in ("OST_AreaSchemeLines", "OST_AreaBoundaryLines"):
        try:
            ids.add(int(getattr(DB.BuiltInCategory, name)))
        except Exception:
            pass
    return ids


def _get_boundary_curves(doc, view):
    curves = []
    boundary_ids = _get_area_boundary_category_ids()
    elements = DB.FilteredElementCollector(doc, view.Id).OfClass(DB.CurveElement)
    for elem in elements:
        cat = elem.Category
        if not cat:
            continue
        if boundary_ids:
            try:
                if cat.Id.IntegerValue not in boundary_ids:
                    continue
            except Exception:
                continue
        else:
            name = cat.Name or ""
            if "Area" not in name or "Boundary" not in name:
                continue
        try:
            curve = elem.GeometryCurve
        except Exception:
            curve = None
        if curve is None:
            continue
        curves.append(curve)
    return curves


def _ensure_sketch_plane(doc, view):
    try:
        if view.SketchPlane:
            return view.SketchPlane
    except Exception:
        pass
    try:
        level = view.GenLevel
    except Exception:
        level = None
    elevation = level.Elevation if level else 0
    plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elevation))
    return DB.SketchPlane.Create(doc, plane)


def _get_area_location_uv(area, view):
    try:
        loc = area.Location
    except Exception:
        loc = None
    if isinstance(loc, DB.LocationPoint):
        pt = loc.Point
        return DB.UV(pt.X, pt.Y)
    try:
        bbox = area.get_BoundingBox(view)
    except Exception:
        bbox = None
    if bbox is None:
        try:
            bbox = area.get_BoundingBox(None)
        except Exception:
            bbox = None
    if bbox is None:
        return None
    center = (bbox.Min + bbox.Max) / 2
    return DB.UV(center.X, center.Y)


def _get_param_value(param):
    if param is None:
        return None
    stype = param.StorageType
    if stype == DB.StorageType.String:
        try:
            return param.AsString()
        except Exception:
            return None
    if stype == DB.StorageType.Double:
        try:
            return param.AsDouble()
        except Exception:
            return None
    if stype == DB.StorageType.Integer:
        try:
            return param.AsInteger()
        except Exception:
            return None
    if stype == DB.StorageType.ElementId:
        try:
            return param.AsElementId()
        except Exception:
            return None
    return None


def _set_param_value(param, value):
    if param is None or param.IsReadOnly:
        return False
    stype = param.StorageType
    try:
        if stype == DB.StorageType.String:
            param.Set("" if value is None else str(value))
            return True
        if stype == DB.StorageType.Double:
            if value is None:
                return False
            param.Set(float(value))
            return True
        if stype == DB.StorageType.Integer:
            if value is None:
                return False
            param.Set(int(value))
            return True
        if stype == DB.StorageType.ElementId:
            if isinstance(value, DB.ElementId):
                param.Set(value)
                return True
    except Exception:
        return False
    return False


def _copy_parameters(source, target, skip_names):
    copied = 0
    for param in source.Parameters:
        if param is None or param.IsReadOnly:
            continue
        name = None
        try:
            name = param.Definition.Name
        except Exception:
            name = None
        if not name or name in skip_names:
            continue
        target_param = None
        try:
            target_param = target.LookupParameter(name)
        except Exception:
            target_param = None
        if not target_param or target_param.IsReadOnly:
            continue
        value = _get_param_value(param)
        if value is None:
            continue
        if _set_param_value(target_param, value):
            copied += 1
    return copied


def _collect_area_tags(doc, view):
    tags = []
    try:
        tags = list(DB.FilteredElementCollector(doc, view.Id).OfClass(DB.AreaTag))
    except Exception:
        try:
            tags = list(DB.FilteredElementCollector(doc, view.Id).OfClass(DB.SpatialElementTag))
        except Exception:
            tags = []
    return tags


def _get_tag_area_id(tag):
    try:
        if hasattr(tag, "TaggedLocalElementId"):
            return tag.TaggedLocalElementId
    except Exception:
        pass
    try:
        if hasattr(tag, "TaggedElementId"):
            link_id = tag.TaggedElementId
            try:
                return link_id.HostElementId
            except Exception:
                return link_id
    except Exception:
        pass
    try:
        ids = tag.GetTaggedLocalElementIds()
        if ids and len(ids) > 0:
            return ids[0]
    except Exception:
        pass
    return None


def _create_area_tag(doc, view, area, point, source_tag=None):
    new_tag = None
    try:
        uv = DB.UV(point.X, point.Y)
        new_tag = doc.Create.NewAreaTag(view, area, uv)
    except Exception:
        new_tag = None

    if new_tag is None:
        try:
            reference = DB.Reference(area)
            new_tag = DB.SpatialElementTag.Create(
                doc,
                view.Id,
                reference,
                False,
                DB.TagOrientation.Horizontal,
                point,
            )
        except Exception:
            new_tag = None

    if new_tag is not None and source_tag is not None:
        try:
            new_tag.ChangeTypeId(source_tag.GetTypeId())
        except Exception:
            pass
    return new_tag


def _pick_index(ui, items, title, prompt):
    selected = ui.uiUtils_select_indices(
        items,
        title=title,
        prompt=prompt,
        multiselect=False,
        width=520,
        height=420,
    )
    return selected[0] if selected else -1


def _pick_source_target_schemes(ui, scheme_names):
    result = ui.uiUtils_select_sheet_renumber_inputs(
        categories=scheme_names,
        print_sets=scheme_names,
        title="Copy Areas - Schemes",
        category_label="Source area scheme:",
        printset_label="Target area scheme:",
        starting_label="(Optional note)",
        cancel_text="Cancel",
        width=620,
        height=340,
    )
    if not result:
        return -1, -1
    source_name = (result.get("category") or "").strip()
    target_name = (result.get("printset") or "").strip()
    if not source_name or not target_name:
        return -1, -1
    try:
        source_index = scheme_names.index(source_name)
        target_index = scheme_names.index(target_name)
    except Exception:
        return -1, -1
    return source_index, target_index


def _pick_levels(ui, items, title, prompt):
    selected = ui.uiUtils_select_indices(
        items,
        title=title,
        prompt=prompt,
        multiselect=True,
        width=720,
        height=620,
    )
    return selected or []


def _pick_copy_tags(ui):
    options = [
        "Copy areas + boundaries only",
        "Copy areas + boundaries + tags",
    ]
    selected = ui.uiUtils_select_indices(
        options,
        title="Copy Areas - Options",
        prompt="Choose what to copy:",
        multiselect=False,
        width=560,
        height=320,
    )
    if not selected:
        return None
    return selected[0] == 1


def _pick_copy_color_scheme(ui):
    options = [
        "Do not copy area color schemes",
        "Copy area color schemes (source -> target views)",
    ]
    selected = ui.uiUtils_select_indices(
        options,
        title="Copy Areas - Options",
        prompt="Include color scheme?",
        multiselect=False,
        width=620,
        height=320,
    )
    if not selected:
        return None
    return selected[0] == 1


def _collect_color_fill_schemes(doc):
    try:
        return list(DB.FilteredElementCollector(doc).OfClass(DB.ColorFillScheme).ToElements())
    except Exception:
        return []


def _get_area_scheme_id_from_color_scheme(scheme):
    try:
        if hasattr(scheme, "AreaSchemeId"):
            return scheme.AreaSchemeId
    except Exception:
        pass
    try:
        method = getattr(scheme, "GetAreaSchemeId", None)
        if callable(method):
            return method()
    except Exception:
        pass
    return None


def _copy_color_fill_scheme_data(source, target):
    for attr in ("Title", "IsByRange", "IsByValue", "IsByPercentage"):
        if hasattr(source, attr) and hasattr(target, attr):
            try:
                setattr(target, attr, getattr(source, attr))
            except Exception:
                pass

    try:
        source_entries = list(source.GetEntries())
    except Exception:
        return False

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

    try:
        set_entries = getattr(target, "SetEntries", None)
        if callable(set_entries):
            set_entries(source_entries)
            return True
    except Exception:
        pass

    add_entry = getattr(target, "AddEntry", None)
    if not callable(add_entry):
        return False
    for entry in source_entries:
        try:
            clone = getattr(entry, "Clone", None)
            new_entry = clone() if callable(clone) else entry
            add_entry(new_entry)
        except Exception:
            continue
    return True


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


def _find_matching_target_color_scheme(doc, source_scheme, target_area_scheme_id):
    source_name = getattr(source_scheme, "Name", "") or ""
    source_category_id = getattr(source_scheme, "CategoryId", None)
    for scheme in _collect_color_fill_schemes(doc):
        try:
            if getattr(scheme, "Name", "") != source_name:
                continue
            if source_category_id and getattr(scheme, "CategoryId", None) != source_category_id:
                continue
            scheme_area_scheme_id = _get_area_scheme_id_from_color_scheme(scheme)
            if scheme_area_scheme_id != target_area_scheme_id:
                continue
            return scheme
        except Exception:
            continue
    return None


def _copy_view_area_color_scheme(doc, source_view, target_view, target_area_scheme_id):
    area_category_id = DB.ElementId(DB.BuiltInCategory.OST_Areas)
    source_scheme_id = _get_view_color_fill_scheme_id(source_view, area_category_id)
    if source_scheme_id is None or source_scheme_id == DB.ElementId.InvalidElementId:
        return False, "source view has no area color scheme assignment"
    source_scheme = doc.GetElement(source_scheme_id)
    if source_scheme is None:
        return False, "source color scheme element not found"

    target_scheme = _find_matching_target_color_scheme(doc, source_scheme, target_area_scheme_id)
    if target_scheme is None:
        return False, "no matching target color scheme found (same name/category in target area scheme)"

    if not _copy_color_fill_scheme_data(source_scheme, target_scheme):
        return False, "failed to copy color scheme entries"
    if not _set_view_color_fill_scheme_id(target_view, area_category_id, target_scheme.Id):
        return False, "failed to assign target color scheme to target view"
    return True, ""


def main():
    ui = _load_uiutils()
    doc = revit.doc
    if doc is None:
        ui.uiUtils_alert("No active document.", title="Copy Areas")
        return

    schemes = _get_area_schemes(doc)
    if not schemes:
        ui.uiUtils_alert("No area schemes found in this project.", title="Copy Areas")
        return

    levels = _get_levels(doc)
    if not levels:
        ui.uiUtils_alert("No levels found in this project.", title="Copy Areas")
        return

    scheme_names = [s.Name for s in schemes]
    level_names = [l.Name for l in levels]

    source_index, target_index = _pick_source_target_schemes(ui, scheme_names)
    if source_index < 0 or target_index < 0:
        return

    if source_index == target_index:
        ui.uiUtils_alert("Source and target area schemes must be different.", title="Copy Areas")
        return

    level_indices = _pick_levels(ui, level_names, "Copy Areas - Levels", "Levels to copy:")
    if not level_indices:
        ui.uiUtils_alert("Select at least one level.", title="Copy Areas")
        return
    copy_tags = _pick_copy_tags(ui)
    if copy_tags is None:
        return
    copy_color_scheme = _pick_copy_color_scheme(ui)
    if copy_color_scheme is None:
        return

    source_scheme = schemes[source_index]
    target_scheme = schemes[target_index]
    selected_levels = [levels[i] for i in level_indices if 0 <= i < len(levels)]

    source_plans = _get_area_plans_by_level(doc, source_scheme.Id)
    target_plans = _get_area_plans_by_level(doc, target_scheme.Id)

    report_lines = []
    total_areas = 0
    total_boundaries = 0
    total_tags = 0
    total_color_scheme_views = 0
    total_failed = 0

    transaction = DB.Transaction(doc, "Copy Areas Between Schemes")
    try:
        transaction.Start()
        for level in selected_levels:
            level_id = level.Id
            source_view = source_plans.get(level_id)
            if not source_view:
                report_lines.append("{}: no source area plan found".format(level.Name))
                continue

            target_view = target_plans.get(level_id)
            if not target_view:
                try:
                    target_view = DB.ViewPlan.CreateAreaPlan(doc, target_scheme.Id, level_id)
                    target_plans[level_id] = target_view
                except Exception as ex:
                    report_lines.append("{}: failed to create target area plan ({})".format(level.Name, ex))
                    total_failed += 1
                    continue

            color_scheme_note = ""
            if copy_color_scheme:
                ok_scheme, scheme_msg = _copy_view_area_color_scheme(
                    doc, source_view, target_view, target_scheme.Id
                )
                if ok_scheme:
                    total_color_scheme_views += 1
                else:
                    total_failed += 1
                    color_scheme_note = " | color scheme: {}".format(scheme_msg)

            boundary_curves = _get_boundary_curves(doc, source_view)
            if boundary_curves:
                sketch_plane = _ensure_sketch_plane(doc, target_view)
                for curve in boundary_curves:
                    try:
                        doc.Create.NewAreaBoundaryLine(sketch_plane, curve, target_view)
                        total_boundaries += 1
                    except Exception:
                        total_failed += 1

            areas = []
            for elem in DB.FilteredElementCollector(doc).OfClass(DB.SpatialElement):
                if not isinstance(elem, DB.Area):
                    continue
                area = elem
                try:
                    if area.AreaScheme.Id != source_scheme.Id:
                        continue
                except Exception:
                    continue
                try:
                    if area.LevelId != level_id:
                        continue
                except Exception:
                    continue
                areas.append(area)

            tag_map = {}
            if copy_tags:
                for tag in _collect_area_tags(doc, source_view):
                    area_id = _get_tag_area_id(tag)
                    if area_id is None:
                        continue
                    tag_map.setdefault(area_id, []).append(tag)

            for area in areas:
                uv = _get_area_location_uv(area, source_view)
                if uv is None:
                    total_failed += 1
                    continue
                try:
                    new_area = doc.Create.NewArea(target_view, uv)
                except Exception:
                    total_failed += 1
                    continue

                try:
                    new_area.Name = area.Name
                except Exception:
                    pass
                try:
                    if hasattr(area, "Number"):
                        new_area.Number = area.Number
                except Exception:
                    pass

                skip_names = {"Area", "Perimeter", "Level", "Name", "Number"}
                _copy_parameters(area, new_area, skip_names)
                total_areas += 1

                if copy_tags:
                    tags = tag_map.get(area.Id, [])
                    for tag in tags:
                        try:
                            point = tag.TagHeadPosition
                        except Exception:
                            point = None
                        if point is None:
                            continue
                        new_tag = _create_area_tag(doc, target_view, new_area, point, source_tag=tag)
                        if new_tag is not None:
                            total_tags += 1
                        else:
                            total_failed += 1

            if copy_tags:
                report_lines.append("{}: areas {} | boundaries {} | tags {}{}".format(
                    level.Name,
                    len(areas),
                    len(boundary_curves),
                    len(tag_map),
                    color_scheme_note,
                ))
            else:
                report_lines.append("{}: areas {} | boundaries {}{}".format(
                    level.Name,
                    len(areas),
                    len(boundary_curves),
                    color_scheme_note,
                ))

        transaction.Commit()
    except Exception:
        try:
            transaction.RollBack()
        except Exception:
            pass
        ui.uiUtils_alert(traceback.format_exc(), title="Copy Areas")
        return

    report_lines.append("")
    report_lines.append("Created areas: {}".format(total_areas))
    report_lines.append("Created boundaries: {}".format(total_boundaries))
    report_lines.append("Created tags: {}".format(total_tags if copy_tags else 0))
    report_lines.append("Applied color schemes to target views: {}".format(
        total_color_scheme_views if copy_color_scheme else 0
    ))
    report_lines.append("Failures: {}".format(total_failed))

    ui.uiUtils_show_text_report(
        "Copy Areas - Results",
        "\n".join(report_lines),
        ok_text="Close",
        cancel_text=None,
        width=780,
        height=620,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui = _load_uiutils()
        ui.uiUtils_alert(traceback.format_exc(), title="Copy Areas")
