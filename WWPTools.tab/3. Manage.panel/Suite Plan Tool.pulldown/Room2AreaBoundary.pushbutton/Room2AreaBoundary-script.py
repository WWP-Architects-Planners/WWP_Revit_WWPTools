import os
import sys
import traceback

from pyrevit import DB, revit
from System.Collections.Generic import List


script_dir = os.path.dirname(__file__)


def _find_lib_path(start_dir):
    current = os.path.abspath(start_dir)
    for _ in range(8):
        candidate = os.path.join(current, "lib")
        if os.path.isdir(candidate):
            return candidate
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    return os.path.abspath(os.path.join(start_dir, "..", "..", "..", "lib"))


lib_path = _find_lib_path(script_dir)
if lib_path not in sys.path:
    sys.path.append(lib_path)




def _elem_id_int(eid):
    try:
        return int(eid.Value)      # Revit 2024+
    except AttributeError:
        return int(eid.Value)  # Revit 2023-

def _load_uiutils():
    import WWP_uiUtils as ui
    return ui


def _get_selected_rooms(doc):
    rooms = []
    try:
        selection = revit.get_selection()
        elems = list(selection.elements)
    except Exception:
        elems = []

    for e in elems:
        if e is None:
            continue
        try:
            if e.Category and _elem_id_int(e.Category.Id) == int(DB.BuiltInCategory.OST_Rooms):
                rooms.append(e)
        except Exception:
            continue

    rooms = [r for r in rooms if _is_valid_room(r)]
    rooms.sort(key=lambda r: ((getattr(r, "Number", "") or ""), (getattr(r, "Name", "") or "")))
    return rooms


def _collect_rooms_on_level(doc, level_id):
    rooms = []
    collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms)
    for room in collector.WhereElementIsNotElementType():
        if room is None:
            continue
        if not _is_valid_room(room):
            continue
        try:
            if room.LevelId != level_id:
                continue
        except Exception:
            continue
        rooms.append(room)
    rooms.sort(key=lambda r: ((getattr(r, "Number", "") or ""), (getattr(r, "Name", "") or "")))
    return rooms


def _is_valid_room(room):
    if room is None:
        return False
    try:
        if room.Area <= 0:
            return False
    except Exception:
        pass
    return True


def _get_area_plan_views(doc):
    views = []
    for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan):
        if v is None:
            continue
        try:
            if v.IsTemplate:
                continue
        except Exception:
            pass
        try:
            if v.ViewType != DB.ViewType.AreaPlan:
                continue
        except Exception:
            continue
        views.append(v)

    def _sort_key(view):
        lvl_name = ""
        try:
            lvl = view.GenLevel
            lvl_name = lvl.Name if lvl else ""
        except Exception:
            lvl_name = ""
        return (lvl_name, view.Name or "")

    views.sort(key=_sort_key)
    return views


def _get_area_scheme_name(doc, view):
    scheme_id = None
    try:
        if hasattr(view, "AreaSchemeId"):
            scheme_id = view.AreaSchemeId
    except Exception:
        scheme_id = None

    if scheme_id is None:
        try:
            if hasattr(view, "AreaScheme") and view.AreaScheme:
                scheme_id = view.AreaScheme.Id
        except Exception:
            scheme_id = None

    if scheme_id is None:
        try:
            p = view.get_Parameter(DB.BuiltInParameter.VIEW_AREA_SCHEME)
            if p:
                scheme_id = p.AsElementId()
        except Exception:
            scheme_id = None

    if scheme_id is None:
        return ""

    try:
        scheme = doc.GetElement(scheme_id)
        return (scheme.Name or "").strip() if scheme else ""
    except Exception:
        return ""


def _format_area_plan_label(doc, view):
    lvl_name = ""
    try:
        lvl = view.GenLevel
        lvl_name = lvl.Name if lvl else ""
    except Exception:
        lvl_name = ""

    scheme_name = _get_area_scheme_name(doc, view)
    view_name = view.Name or ""

    if scheme_name:
        return "{} - {} - {}".format(scheme_name, lvl_name, view_name)
    return "{} - {}".format(lvl_name, view_name)


def _get_area_boundary_category_ids():
    ids = set()
    for name in ("OST_AreaSchemeLines", "OST_AreaBoundaryLines"):
        try:
            ids.add(int(getattr(DB.BuiltInCategory, name)))
        except Exception:
            pass
    return ids


def _round_coord(value, step):
    try:
        return int(round(float(value) / float(step)))
    except Exception:
        return 0


def _round_xyz(pt, step):
    return (
        _round_coord(getattr(pt, "X", 0.0), step),
        _round_coord(getattr(pt, "Y", 0.0), step),
        _round_coord(getattr(pt, "Z", 0.0), step),
    )


def _curve_key(curve, step=1e-4):
    """Stable-ish key for overlap detection.

    Uses endpoints (order-independent) + midpoint sampled at normalized 0.5.
    step is in Revit internal units (feet).
    """
    if curve is None:
        return None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return None

    try:
        pm = curve.Evaluate(0.5, True)
    except Exception:
        try:
            pm = (p0 + p1) / 2
        except Exception:
            pm = p0

    a = _round_xyz(p0, step)
    b = _round_xyz(p1, step)
    m = _round_xyz(pm, step)
    ends = (a, b) if a <= b else (b, a)
    try:
        ctype = curve.GetType().FullName
    except Exception:
        ctype = "Curve"
    return (ctype, ends[0], ends[1], m)


def _delete_duplicate_boundaries_in_view(doc, view, step=1e-4):
    boundary_ids = _get_area_boundary_category_ids()
    seen = set()
    dup_ids = []

    collector = DB.FilteredElementCollector(doc, view.Id).OfClass(DB.CurveElement)
    for elem in collector:
        if elem is None:
            continue
        try:
            cat = elem.Category
            if not cat:
                continue
            if boundary_ids:
                if _elem_id_int(cat.Id) not in boundary_ids:
                    continue
            else:
                cname = (cat.Name or "").lower()
                if "area" not in cname or "bound" not in cname:
                    continue
        except Exception:
            continue

        try:
            curve = elem.GeometryCurve
        except Exception:
            curve = None
        key = _curve_key(curve, step=step)
        if key is None:
            continue
        if key in seen:
            dup_ids.append(elem.Id)
        else:
            seen.add(key)

    if not dup_ids:
        return 0, seen

    try:
        net_ids = List[DB.ElementId]()
        for eid in dup_ids:
            net_ids.Add(eid)
        doc.Delete(net_ids)
    except Exception:
        # If delete fails, just keep going; caller still gets current seen keys.
        pass

    return len(dup_ids), seen


def _ensure_sketch_plane(doc, view):
    try:
        if view.SketchPlane:
            return view.SketchPlane
    except Exception:
        pass
    elevation = 0.0
    try:
        lvl = view.GenLevel
        elevation = lvl.Elevation if lvl else 0.0
    except Exception:
        elevation = 0.0
    plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elevation))
    return DB.SketchPlane.Create(doc, plane)


def _get_room_boundary_curves(room):
    opts = DB.SpatialElementBoundaryOptions()
    try:
        opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
    except Exception:
        pass

    curves = []
    try:
        loops = room.GetBoundarySegments(opts)
    except Exception:
        loops = None

    if not loops:
        return curves

    for loop in loops:
        if not loop:
            continue
        for seg in loop:
            if seg is None:
                continue
            try:
                c = seg.GetCurve()
            except Exception:
                c = None
            if c is None:
                continue
            curves.append(c)
    return curves


def _move_curve_to_elevation(curve, target_elevation):
    if curve is None:
        return None
    try:
        z0 = curve.GetEndPoint(0).Z
    except Exception:
        z0 = 0.0
    dz = float(target_elevation) - float(z0)
    try:
        xf = DB.Transform.CreateTranslation(DB.XYZ(0, 0, dz))
        return curve.CreateTransformed(xf)
    except Exception:
        return curve


def _get_location_uv(elem):
    try:
        loc = elem.Location
    except Exception:
        loc = None
    if isinstance(loc, DB.LocationPoint):
        pt = loc.Point
        return DB.UV(pt.X, pt.Y)
    if isinstance(loc, DB.LocationCurve):
        try:
            pt = loc.Curve.Evaluate(0.5, True)
            return DB.UV(pt.X, pt.Y)
        except Exception:
            return None
    return None


def _get_param_guid(param):
    try:
        return param.GUID
    except Exception:
        return None


def _find_target_param_by_guid(target_elem, guid):
    if target_elem is None or guid is None:
        return None
    for p in target_elem.Parameters:
        if p is None:
            continue
        try:
            if p.GUID == guid:
                return p
        except Exception:
            continue
    return None


def _get_param_name(param):
    try:
        return param.Definition.Name
    except Exception:
        return None


def _get_param_value(param):
    if param is None:
        return None
    stype = param.StorageType
    try:
        if stype == DB.StorageType.String:
            return param.AsString()
        if stype == DB.StorageType.Double:
            return param.AsDouble()
        if stype == DB.StorageType.Integer:
            return param.AsInteger()
        if stype == DB.StorageType.ElementId:
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
            if value is None:
                return False
            param.Set(str(value))
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


def _copy_room_params_to_area(room, area, skip_names=None):
    if skip_names is None:
        skip_names = set()
    copied = 0
    failed = 0

    for src in room.Parameters:
        if src is None:
            continue
        try:
            if src.IsReadOnly:
                continue
        except Exception:
            continue

        name = _get_param_name(src)
        if not name or name in skip_names:
            continue

        value = _get_param_value(src)
        if value is None:
            continue

        tgt = None
        guid = _get_param_guid(src)
        if guid is not None:
            tgt = _find_target_param_by_guid(area, guid)

        if tgt is None:
            try:
                tgt = area.LookupParameter(name)
            except Exception:
                tgt = None

        if tgt is None:
            continue

        try:
            if tgt.IsReadOnly:
                continue
        except Exception:
            continue

        try:
            if tgt.StorageType != src.StorageType:
                continue
        except Exception:
            continue

        if _set_param_value(tgt, value):
            copied += 1
        else:
            failed += 1

    return copied, failed


def _get_room_name_number(room):
    """Return (name, number) using Revit's built-in room parameters."""
    name = None
    number = None

    try:
        p_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
        if p_name:
            name = p_name.AsString()
    except Exception:
        name = None

    try:
        p_num = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER)
        if p_num:
            number = p_num.AsString()
    except Exception:
        number = None

    if name is None:
        try:
            name = getattr(room, "Name", None)
        except Exception:
            name = None
    if number is None:
        try:
            number = getattr(room, "Number", None)
        except Exception:
            number = None

    return (name or ""), (number or "")


def main():
    ui = _load_uiutils()
    doc = revit.doc
    if doc is None:
        ui.uiUtils_alert("No active document.", title="Room to Area Boundaries")
        return

    area_plans = _get_area_plan_views(doc)
    if not area_plans:
        ui.uiUtils_alert("No Area Plan views found in this project.", title="Room to Area Boundaries")
        return

    labels = [_format_area_plan_label(doc, v) for v in area_plans]
    selected_indices = ui.uiUtils_select_indices(
        labels,
        title="Room to Area Boundaries",
        prompt="Select target Area Plan(s) (levels):",
        multiselect=True,
        width=980,
        height=540,
    )
    if not selected_indices:
        return

    target_views = []
    for idx in selected_indices:
        try:
            i = int(idx)
        except Exception:
            continue
        if 0 <= i < len(area_plans):
            target_views.append(area_plans[i])

    if not target_views:
        return

    # Determine room set once (then filter per level)
    selected_rooms = _get_selected_rooms(doc)
    if not selected_rooms:
        if not ui.uiUtils_confirm(
            "No Rooms selected.\nProcess ALL Rooms on the selected level(s)?",
            title="Room to Area Boundaries",
        ):
            return
    
    # Totals across all targets
    total_created_boundaries = 0
    total_failed_boundaries = 0
    total_skipped_overlapped = 0
    total_deleted_overlapped = 0
    total_created_areas = 0
    total_failed_areas = 0
    total_copied_params = 0
    total_failed_params = 0
    total_rooms_processed = 0

    report = []

    # Avoid copying these since Name/Number are handled directly
    skip_names = {"Name", "Number", "Area", "Perimeter", "Level"}

    t = DB.Transaction(doc, "Room Boundary to Area Boundary")
    try:
        t.Start()

        for target_view in target_views:
            target_level = None
            try:
                target_level = target_view.GenLevel
            except Exception:
                target_level = None
            if target_level is None:
                report.append("{}: skipped (no associated level)".format(target_view.Name))
                continue

            elevation = 0.0
            try:
                elevation = float(target_level.Elevation)
            except Exception:
                elevation = 0.0

            rooms = []
            if selected_rooms:
                rooms = [r for r in selected_rooms if getattr(r, "LevelId", None) == target_level.Id and _is_valid_room(r)]
            else:
                rooms = _collect_rooms_on_level(doc, target_level.Id)

            if not rooms:
                report.append("{}: {} | rooms 0 (nothing to do)".format(
                    _format_area_plan_label(doc, target_view),
                    target_level.Name,
                ))
                continue

            sketch_plane = _ensure_sketch_plane(doc, target_view)

            created_boundaries = 0
            failed_boundaries = 0
            skipped_overlapped_boundaries = 0
            deleted_overlapped_boundaries = 0
            created_areas = 0
            failed_areas = 0
            copied_params = 0
            failed_params = 0

            deleted_overlapped_boundaries, existing_keys = _delete_duplicate_boundaries_in_view(
                doc,
                target_view,
                step=1e-4,
            )
            if existing_keys is None:
                existing_keys = set()

            for room in rooms:
                curves = _get_room_boundary_curves(room)
                for c in curves:
                    tc = _move_curve_to_elevation(c, elevation)
                    key = _curve_key(tc, step=1e-4)
                    if key is not None and key in existing_keys:
                        skipped_overlapped_boundaries += 1
                        continue
                    try:
                        doc.Create.NewAreaBoundaryLine(sketch_plane, tc, target_view)
                        created_boundaries += 1
                        if key is not None:
                            existing_keys.add(key)
                    except Exception:
                        failed_boundaries += 1

                uv = _get_location_uv(room)
                if uv is None:
                    failed_areas += 1
                    continue

                try:
                    area = doc.Create.NewArea(target_view, uv)
                except Exception:
                    failed_areas += 1
                    continue

                created_areas += 1

                room_name, room_number = _get_room_name_number(room)
                try:
                    if room_number:
                        area.Number = room_number
                except Exception:
                    pass
                try:
                    if room_name:
                        area.Name = room_name
                except Exception:
                    pass

                c_ok, c_fail = _copy_room_params_to_area(room, area, skip_names=skip_names)
                copied_params += c_ok
                failed_params += c_fail

            report.append(
                "{} | rooms {} | areas {} | boundaries {} | skipped dup {} | deleted dup {}".format(
                    _format_area_plan_label(doc, target_view),
                    len(rooms),
                    created_areas,
                    created_boundaries,
                    skipped_overlapped_boundaries,
                    deleted_overlapped_boundaries,
                )
            )

            total_rooms_processed += len(rooms)
            total_created_boundaries += created_boundaries
            total_failed_boundaries += failed_boundaries
            total_skipped_overlapped += skipped_overlapped_boundaries
            total_deleted_overlapped += deleted_overlapped_boundaries
            total_created_areas += created_areas
            total_failed_areas += failed_areas
            total_copied_params += copied_params
            total_failed_params += failed_params

        t.Commit()

    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        ui.uiUtils_alert(traceback.format_exc(), title="Room to Area Boundaries")
        return

    report.append("")
    report.append("Totals")
    report.append("Targets processed: {}".format(len(target_views)))
    report.append("Rooms processed: {}".format(total_rooms_processed))
    report.append("Created Area Boundary Lines: {}".format(total_created_boundaries))
    report.append("Failed Area Boundary Lines: {}".format(total_failed_boundaries))
    report.append("Skipped overlapped boundary curves: {}".format(total_skipped_overlapped))
    report.append("Deleted overlapped existing boundaries: {}".format(total_deleted_overlapped))
    report.append("Created Areas: {}".format(total_created_areas))
    report.append("Failed Areas: {}".format(total_failed_areas))
    report.append("Copied parameter values: {}".format(total_copied_params))
    report.append("Failed parameter sets (writable match but Set failed): {}".format(total_failed_params))

    ui.uiUtils_show_text_report(
        "Room to Area Boundaries - Results",
        "\n".join(report),
        ok_text="Close",
        cancel_text=None,
        width=900,
        height=620,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception:
        try:
            ui = _load_uiutils()
            ui.uiUtils_alert(traceback.format_exc(), title="Room to Area Boundaries")
        except Exception:
            raise
