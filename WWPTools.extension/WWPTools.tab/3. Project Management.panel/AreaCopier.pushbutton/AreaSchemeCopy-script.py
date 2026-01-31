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
        try:
            param = view.get_Parameter(DB.BuiltInParameter.VIEW_AREA_SCHEME)
            if not param or param.AsElementId() != scheme_id:
                continue
        except Exception:
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

    source_index = _pick_index(ui, scheme_names, "Copy Areas - Source", "Source area scheme:")
    if source_index < 0:
        return

    target_index = _pick_index(ui, scheme_names, "Copy Areas - Target", "Target area scheme:")
    if target_index < 0:
        return

    if source_index == target_index:
        ui.uiUtils_alert("Source and target area schemes must be different.", title="Copy Areas")
        return

    level_indices = _pick_levels(ui, level_names, "Copy Areas - Levels", "Levels to copy:")
    if not level_indices:
        ui.uiUtils_alert("Select at least one level.", title="Copy Areas")
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
            for area in DB.FilteredElementCollector(doc).OfClass(DB.Area):
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

            report_lines.append("{}: areas {} | boundaries {} | tags {}".format(
                level.Name,
                len(areas),
                len(boundary_curves),
                len(tag_map)
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
    report_lines.append("Created tags: {}".format(total_tags))
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
