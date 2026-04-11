import csv
import math
import os
import re
import sys
import traceback

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit import UI
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, DB


TITLE = "Fix Floor Heights from CSV"
DEFAULT_UNIT_TEXT = "meters"
DEFAULT_TOLERANCE_TEXT = "0.025"
MATCH_EPSILON = 1e-6


SCRIPT_DIR = os.path.dirname(__file__)
LIB_PATH = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "..", "..", "lib"))
if LIB_PATH not in sys.path:
    sys.path.append(LIB_PATH)

import WWP_uiUtils as ui
from System.Collections.Generic import List
from WWP_compat import io_open


def _elem_id_int(eid):
    try:
        return int(eid.Value)      # Revit 2024+
    except AttributeError:
        return int(eid.Value)  # Revit 2023-


HEADER_ALIASES = {
    "x": set([
        "x",
        "coordx",
        "coordinatex",
        "centerx",
        "east",
        "easting",
        "globalx",
        "localx",
        "pointx",
        "xcoord",
        "xcoordinate",
        "xpos",
    ]),
    "y": set([
        "y",
        "coordy",
        "coordinatey",
        "centery",
        "globaly",
        "localy",
        "north",
        "northing",
        "pointy",
        "ycoord",
        "ycoordinate",
        "ypos",
    ]),
    "z": set([
        "altitude",
        "centerz",
        "coordz",
        "coordinatez",
        "elev",
        "elevation",
        "globalz",
        "height",
        "localz",
        "pointz",
        "reducedlevel",
        "rl",
        "z",
        "zcoord",
        "zcoordinate",
        "zpos",
    ]),
}


class CsvPoint(object):
    def __init__(self, x, y, z, row_number):
        self.x = x
        self.y = y
        self.z = z
        self.row_number = row_number


def _element_id_value(elem_id):
    if elem_id is None:
        return None
    if hasattr(elem_id, "IntegerValue"):
        return _elem_id_int(elem_id)
    if hasattr(elem_id, "Value"):
        return elem_id.Value
    try:
        return int(elem_id)
    except Exception:
        return None


def _is_floor(element):
    if element is None:
        return False
    try:
        if isinstance(element, DB.Floor):
            return True
    except Exception:
        pass
    try:
        category = element.Category
        return category is not None and _elem_id_int(category.Id) == int(DB.BuiltInCategory.OST_Floors)
    except Exception:
        return False


def _default_directory(doc):
    try:
        if doc and doc.PathName:
            folder = os.path.dirname(doc.PathName)
            if os.path.isdir(folder):
                return folder
    except Exception:
        pass
    return os.path.expanduser("~")


def _normalize_header(value):
    return re.sub(r"[^a-z0-9]", "", (value or "").strip().lower())


def _parse_float(value):
    text = (value or "").strip()
    if not text:
        return None
    text = text.replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(",", "")
    try:
        return float(text)
    except Exception:
        return None


def _get_unit_id(unit_name):
    name = (unit_name or "").strip().lower()
    if name in ("ft", "foot", "feet"):
        return getattr(DB.UnitTypeId, "Feet", None) or DB.DisplayUnitType.DUT_DECIMAL_FEET
    if name in ("in", "inch", "inches"):
        return getattr(DB.UnitTypeId, "Inches", None) or DB.DisplayUnitType.DUT_DECIMAL_INCHES
    if name in ("m", "meter", "meters", "metre", "metres"):
        return getattr(DB.UnitTypeId, "Meters", None) or DB.DisplayUnitType.DUT_METERS
    if name in ("cm", "centimeter", "centimeters", "centimetre", "centimetres"):
        return getattr(DB.UnitTypeId, "Centimeters", None) or DB.DisplayUnitType.DUT_CENTIMETERS
    if name in ("mm", "millimeter", "millimeters", "millimetre", "millimetres"):
        return getattr(DB.UnitTypeId, "Millimeters", None) or DB.DisplayUnitType.DUT_MILLIMETERS
    return None


def _convert_to_internal(value, unit_name):
    unit_id = _get_unit_id(unit_name)
    if unit_id is None:
        raise ValueError("Unsupported unit '{}'".format(unit_name))
    return DB.UnitUtils.ConvertToInternalUnits(float(value), unit_id)


def _pick_floors(doc, uidoc):
    selected = []
    try:
        for element_id in uidoc.Selection.GetElementIds():
            element = doc.GetElement(element_id)
            if _is_floor(element):
                selected.append(element)
    except Exception:
        pass

    if not selected:
        try:
            for element in revit.get_selection().elements:
                if _is_floor(element):
                    selected.append(element)
        except Exception:
            pass

    if selected:
        return selected

    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            "Select floor(s) to update from CSV",
        )
    except OperationCanceledException:
        return []

    floors = []
    non_floor_labels = []
    for ref in refs:
        try:
            element = doc.GetElement(ref)
            if element is None and hasattr(ref, "ElementId"):
                element = doc.GetElement(ref.ElementId)
            if _is_floor(element):
                floors.append(element)
            elif element is not None:
                try:
                    cat_name = element.Category.Name if element.Category else type(element).__name__
                except Exception:
                    cat_name = type(element).__name__
                non_floor_labels.append(cat_name)
        except Exception:
            continue
    if not floors and non_floor_labels:
        ui.uiUtils_alert(
            "Picked elements were not recognized as floors.\n\nFirst categories found:\n- {}".format(
                "\n- ".join(non_floor_labels[:10])
            ),
            title=TITLE,
        )
    return floors


def _prompt_csv_file(doc):
    return ui.uiUtils_open_file_dialog(
        title=TITLE,
        filter_text="CSV files (*.csv)|*.csv|All files (*.*)|*.*",
        multiselect=False,
        initial_directory=_default_directory(doc),
    )


def _prompt_settings():
    unit_text = ui.uiUtils_prompt_text(
        title=TITLE,
        prompt="CSV units for X, Y, and Z (feet, inches, meters, centimeters, millimeters):",
        default_value=DEFAULT_UNIT_TEXT,
        ok_text="Next",
        cancel_text="Cancel",
        width=520,
        height=220,
    )
    if unit_text is None:
        return None

    tolerance_text = ui.uiUtils_prompt_text(
        title=TITLE,
        prompt="XY match tolerance in the same CSV units:",
        default_value=DEFAULT_TOLERANCE_TEXT,
        ok_text="Continue",
        cancel_text="Cancel",
        width=420,
        height=220,
    )
    if tolerance_text is None:
        return None

    unit_name = (unit_text or "").strip().lower()
    if _get_unit_id(unit_name) is None:
        ui.uiUtils_alert(
            "Unsupported units '{}'. Use feet, inches, meters, centimeters, or millimeters.".format(unit_text),
            title=TITLE,
        )
        return None

    tolerance_value = _parse_float(tolerance_text)
    if tolerance_value is None or tolerance_value <= 0:
        ui.uiUtils_alert("Tolerance must be a positive number.", title=TITLE)
        return None

    return {
        "unit_name": unit_name,
        "tolerance_csv": tolerance_value,
        "tolerance_internal": _convert_to_internal(tolerance_value, unit_name),
    }


def _sniff_dialect(sample_text):
    try:
        return csv.Sniffer().sniff(sample_text, delimiters=",;\t|")
    except Exception:
        return csv.excel


def _infer_column_indexes(rows):
    if not rows:
        raise ValueError("CSV is empty.")

    first_row = rows[0]
    normalized = [_normalize_header(value) for value in first_row]
    indexes = {}
    for axis, aliases in HEADER_ALIASES.items():
        for idx, value in enumerate(normalized):
            if value in aliases:
                indexes[axis] = idx
                break

    if len(indexes) == 3:
        return indexes, rows[1:], True

    numeric_indexes = []
    for idx, value in enumerate(first_row):
        if _parse_float(value) is not None:
            numeric_indexes.append(idx)

    if len(numeric_indexes) >= 3:
        return {
            "x": numeric_indexes[0],
            "y": numeric_indexes[1],
            "z": numeric_indexes[2],
        }, rows, False

    raise ValueError(
        "Could not detect X/Y/Z columns. Use headers like X, Y, Z or Elevation, or place X/Y/Z in the first three numeric columns."
    )


def _read_csv_points(csv_path, unit_name):
    with io_open(csv_path, "r", encoding="utf-8-sig", newline="") as csv_file:
        sample = csv_file.read(4096)
        csv_file.seek(0)
        reader = csv.reader(csv_file, _sniff_dialect(sample))
        rows = [row for row in reader if row and any((cell or "").strip() for cell in row)]

    indexes, data_rows, has_header = _infer_column_indexes(rows)

    points = []
    skipped = []
    for idx, row in enumerate(data_rows, start=2 if has_header else 1):
        try:
            x_value = row[indexes["x"]] if indexes["x"] < len(row) else ""
            y_value = row[indexes["y"]] if indexes["y"] < len(row) else ""
            z_value = row[indexes["z"]] if indexes["z"] < len(row) else ""
            x = _parse_float(x_value)
            y = _parse_float(y_value)
            z = _parse_float(z_value)
            if x is None or y is None or z is None:
                skipped.append("Row {} skipped: invalid X/Y/Z values.".format(idx))
                continue
            points.append(
                CsvPoint(
                    _convert_to_internal(x, unit_name),
                    _convert_to_internal(y, unit_name),
                    _convert_to_internal(z, unit_name),
                    idx,
                )
            )
        except Exception as exc:
            skipped.append("Row {} skipped: {}".format(idx, exc))

    if not points:
        raise ValueError("No valid CSV points were found.")

    return points, skipped, indexes


def _bucket_key(x, y, bucket_size):
    return int(math.floor(x / bucket_size)), int(math.floor(y / bucket_size))


def _build_spatial_index(points, tolerance_internal):
    bucket_size = max(float(tolerance_internal), 1e-9)
    index = {}
    for point in points:
        key = _bucket_key(point.x, point.y, bucket_size)
        index.setdefault(key, []).append(point)
    return index, bucket_size


def _find_best_point(x, y, spatial_index, bucket_size, tolerance_internal):
    tolerance_sq = tolerance_internal * tolerance_internal
    base_x, base_y = _bucket_key(x, y, bucket_size)

    best_point = None
    best_dist_sq = None

    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for point in spatial_index.get((base_x + dx, base_y + dy), []):
                dist_sq = ((point.x - x) ** 2) + ((point.y - y) ** 2)
                if dist_sq > tolerance_sq:
                    continue
                if best_dist_sq is None or dist_sq < best_dist_sq:
                    best_point = point
                    best_dist_sq = dist_sq

    return best_point, math.sqrt(best_dist_sq) if best_dist_sq is not None else None


def _iter_solids(geometry_element):
    if geometry_element is None:
        return
    for geom_obj in geometry_element:
        if isinstance(geom_obj, DB.Solid):
            try:
                if geom_obj.Volume > 0:
                    yield geom_obj
            except Exception:
                continue
        elif isinstance(geom_obj, DB.GeometryInstance):
            for child in _iter_solids(geom_obj.GetInstanceGeometry()):
                yield child


def _get_floor_top_faces(floor):
    options = DB.Options()
    options.ComputeReferences = False
    options.IncludeNonVisibleObjects = False

    try:
        geometry = floor.get_Geometry(options)
    except Exception:
        geometry = None

    faces = []
    for solid in _iter_solids(geometry):
        for face in solid.Faces:
            if not isinstance(face, DB.PlanarFace):
                continue
            try:
                normal = face.FaceNormal
            except Exception:
                continue
            if normal is None or normal.Z <= 0.01:
                continue
            faces.append(face)
    return faces


def _point_is_on_floor_top_face(point, faces, tolerance_internal):
    if not faces:
        return True

    xy_tolerance = max(float(tolerance_internal), 1e-6)
    for face in faces:
        try:
            probe = DB.XYZ(point.x, point.y, face.Origin.Z)
            result = face.Project(probe)
            if result is None or result.XYZPoint is None:
                continue
            xyz_point = result.XYZPoint
            if abs(xyz_point.X - point.x) > xy_tolerance or abs(xyz_point.Y - point.y) > xy_tolerance:
                continue
            uv_point = result.UVPoint
            if uv_point is not None and face.IsInside(uv_point):
                return True
        except Exception:
            continue
    return False


def _points_for_floor_footprint(floor, points, tolerance_internal):
    try:
        bbox = floor.get_BoundingBox(None)
    except Exception:
        bbox = None

    faces = _get_floor_top_faces(floor)

    if bbox is None and not faces:
        return list(points)

    if bbox is None:
        min_x = min_y = float("-inf")
        max_x = max_y = float("inf")
    else:
        min_x = bbox.Min.X - tolerance_internal
        min_y = bbox.Min.Y - tolerance_internal
        max_x = bbox.Max.X + tolerance_internal
        max_y = bbox.Max.Y + tolerance_internal

    filtered = []
    for point in points:
        if not (min_x <= point.x <= max_x and min_y <= point.y <= max_y):
            continue
        if _point_is_on_floor_top_face(point, faces, tolerance_internal):
            filtered.append(point)
    return filtered


def _find_best_vertex(x, y, vertices, tolerance_internal, used_indexes):
    tolerance_sq = tolerance_internal * tolerance_internal
    best_index = None
    best_vertex = None
    best_dist_sq = None

    for idx, vertex in enumerate(vertices):
        if idx in used_indexes:
            continue
        position = vertex.Position
        dist_sq = ((position.X - x) ** 2) + ((position.Y - y) ** 2)
        if dist_sq > tolerance_sq:
            continue
        if best_dist_sq is None or dist_sq < best_dist_sq:
            best_index = idx
            best_vertex = vertex
            best_dist_sq = dist_sq

    if best_vertex is None:
        return None, None, None
    return best_vertex, best_index, math.sqrt(best_dist_sq)


def _floor_label(floor):
    try:
        name = floor.Name or "<unnamed>"
    except Exception:
        name = "<unnamed>"
    return "Floor {} ({})".format(_element_id_value(getattr(floor, "Id", None)), name)


def _get_slab_shape_editor(floor):
    for accessor_name in ("SlabShapeEditor", "get_SlabShapeEditor"):
        try:
            accessor = getattr(floor, accessor_name, None)
            if callable(accessor):
                editor = accessor()
            else:
                editor = accessor
            if editor is not None:
                return editor
        except Exception:
            pass

    try:
        method = getattr(floor, "GetSlabShapeEditor", None)
        if callable(method):
            editor = method()
            if editor is not None:
                return editor
    except Exception:
        pass

    return None


def _add_shape_points(editor, points):
    added = []
    failed = []
    if not points:
        return added, failed

    draw_point = getattr(editor, "DrawPoint", None)
    add_points = getattr(editor, "AddPoints", None)
    if not callable(draw_point) and not callable(add_points):
        raise AttributeError("SlabShapeEditor does not expose DrawPoint or AddPoints")

    for point in points:
        location = DB.XYZ(point.x, point.y, point.z)
        try:
            if callable(draw_point):
                draw_point(location)
            else:
                net_points = List[DB.XYZ]()
                net_points.Add(location)
                add_points(net_points)
            added.append(point)
        except Exception as exc:
            failed.append("Row {} add failed: {}".format(point.row_number, exc))

    return added, failed


def _plan_floor(floor, points, tolerance_internal, enable_editor):
    label = _floor_label(floor)
    plan = {
        "floor": floor,
        "label": label,
        "vertices": 0,
        "candidate_points": 0,
        "matched": 0,
        "updated": 0,
        "added": 0,
        "unmatched_vertices": 0,
        "updates": [],
        "additions": [],
        "add_failures": [],
        "error": None,
    }

    try:
        editor = _get_slab_shape_editor(floor)
        if editor is None:
            plan["error"] = "No SlabShapeEditor available."
            return plan

        if enable_editor and not editor.IsEnabled:
            editor.Enable()

        vertices = list(editor.SlabShapeVertices)
        plan["vertices"] = len(vertices)
        candidate_points = _points_for_floor_footprint(floor, points, tolerance_internal)
        plan["candidate_points"] = len(candidate_points)

        used_vertex_indexes = set()
        for point in candidate_points:
            vertex, vertex_index, distance = _find_best_vertex(
                point.x,
                point.y,
                vertices,
                tolerance_internal,
                used_vertex_indexes,
            )
            if vertex is None:
                plan["added"] += 1
                plan["additions"].append(point)
                continue

            used_vertex_indexes.add(vertex_index)
            delta_z = point.z - vertex.Position.Z
            plan["matched"] += 1
            if abs(delta_z) <= MATCH_EPSILON:
                continue

            plan["updated"] += 1
            plan["updates"].append({
                "vertex": vertex,
                "delta_z": delta_z,
                "distance": distance,
                "csv_row": point.row_number,
            })

        plan["unmatched_vertices"] = max(0, len(vertices) - len(used_vertex_indexes))
    except Exception as exc:
        plan["error"] = str(exc)

    return plan


def _preview_plans(doc, floors, points, tolerance_internal):
    transaction = DB.Transaction(doc, TITLE + " Preview")
    plans = []
    transaction.Start()
    try:
        for floor in floors:
            plans.append(_plan_floor(floor, points, tolerance_internal, enable_editor=True))
    finally:
        transaction.RollBack()
    return plans


def _build_preview_report(csv_path, settings, points, skipped_rows, plans):
    lines = [
        "CSV: {}".format(csv_path),
        "Units: {}".format(settings["unit_name"]),
        "XY tolerance: {} {}".format(settings["tolerance_csv"], settings["unit_name"]),
        "Points loaded: {}".format(len(points)),
        "Rows skipped: {}".format(len(skipped_rows)),
        "Floors selected: {}".format(len(plans)),
        "",
        "Preview:",
    ]

    total_vertices = 0
    total_candidate_points = 0
    total_matched = 0
    total_updated = 0
    total_added = 0
    total_add_failures = 0
    total_unmatched_vertices = 0
    error_count = 0

    for plan in plans:
        total_vertices += plan["vertices"]
        total_candidate_points += plan["candidate_points"]
        total_matched += plan["matched"]
        total_updated += plan["updated"]
        total_added += plan["added"]
        total_add_failures += len(plan["add_failures"])
        total_unmatched_vertices += plan["unmatched_vertices"]
        if plan["error"]:
            error_count += 1
            lines.append("- {}: {}".format(plan["label"], plan["error"]))
        else:
            lines.append(
                "- {}: vertices {}, csv points {}, matched {}, will update {}, will add {}, untouched vertices {}".format(
                    plan["label"],
                    plan["vertices"],
                    plan["candidate_points"],
                    plan["matched"],
                    plan["updated"],
                    plan["added"],
                    plan["unmatched_vertices"],
                )
            )

    lines.extend([
        "",
        "Totals:",
        "- Vertices found: {}".format(total_vertices),
        "- CSV points on selected floor extents: {}".format(total_candidate_points),
        "- Matched by XY: {}".format(total_matched),
        "- Will update Z: {}".format(total_updated),
        "- Will add points: {}".format(total_added),
        "- Add failures after apply: {}".format(total_add_failures),
        "- Existing vertices not touched: {}".format(total_unmatched_vertices),
        "- Floors with errors: {}".format(error_count),
    ])

    if skipped_rows:
        lines.extend(["", "Skipped CSV rows (first 20):"])
        lines.extend("- {}".format(item) for item in skipped_rows[:20])

    return "\n".join(lines), total_updated + total_added


def _apply_updates(doc, floors, points, tolerance_internal):
    transaction = DB.Transaction(doc, TITLE)
    plans = []
    transaction.Start()
    try:
        for floor in floors:
            plan = _plan_floor(floor, points, tolerance_internal, enable_editor=True)
            if not plan["error"]:
                editor = _get_slab_shape_editor(floor)
                for update in plan["updates"]:
                    editor.ModifySubElement(update["vertex"], update["delta_z"])
                added_points, failed_adds = _add_shape_points(editor, plan["additions"])
                plan["added"] = len(added_points)
                plan["add_failures"] = failed_adds
            plans.append(plan)
        transaction.Commit()
        return plans
    except Exception:
        transaction.RollBack()
        raise


def _build_result_report(plans):
    lines = ["Results:"]

    total_vertices = 0
    total_candidate_points = 0
    total_matched = 0
    total_updated = 0
    total_added = 0
    total_add_failures = 0
    total_unmatched_vertices = 0
    total_errors = 0

    for plan in plans:
        total_vertices += plan["vertices"]
        total_candidate_points += plan["candidate_points"]
        total_matched += plan["matched"]
        total_updated += plan["updated"]
        total_added += plan["added"]
        total_add_failures += len(plan["add_failures"])
        total_unmatched_vertices += plan["unmatched_vertices"]
        if plan["error"]:
            total_errors += 1
            lines.append("- {}: {}".format(plan["label"], plan["error"]))
        else:
            lines.append(
                "- {}: updated {}, added {}, matched {}, untouched vertices {}".format(
                    plan["label"],
                    plan["updated"],
                    plan["added"],
                    plan["matched"],
                    plan["unmatched_vertices"],
                )
            )

    lines.extend([
        "",
        "Totals:",
        "- Vertices found: {}".format(total_vertices),
        "- CSV points on selected floor extents: {}".format(total_candidate_points),
        "- Matched by XY: {}".format(total_matched),
        "- Updated Z: {}".format(total_updated),
        "- Added points: {}".format(total_added),
        "- Add failures: {}".format(total_add_failures),
        "- Existing vertices not touched: {}".format(total_unmatched_vertices),
        "- Floors with errors: {}".format(total_errors),
    ])

    failed_rows = []
    for plan in plans:
        failed_rows.extend(plan["add_failures"])
    if failed_rows:
        lines.extend([
            "",
            "Add failures (first 20):",
        ])
        lines.extend("- {}".format(item) for item in failed_rows[:20])

    return "\n".join(lines)


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None or uidoc is None:
        UI.TaskDialog.Show(TITLE, "No active Revit document found.")
        return

    floors = _pick_floors(doc, uidoc)
    if not floors:
        UI.TaskDialog.Show(TITLE, "No floors selected.")
        return

    csv_path = _prompt_csv_file(doc)
    if not csv_path:
        return

    settings = _prompt_settings()
    if not settings:
        return

    try:
        points, skipped_rows, _ = _read_csv_points(csv_path, settings["unit_name"])
        preview_plans = _preview_plans(
            doc,
            floors,
            points,
            settings["tolerance_internal"],
        )
        preview_text, update_count = _build_preview_report(
            csv_path,
            settings,
            points,
            skipped_rows,
            preview_plans,
        )
        if update_count <= 0:
            ui.uiUtils_show_text_report(
                TITLE + " Preview",
                preview_text,
                ok_text="Close",
                cancel_text=None,
                width=860,
                height=620,
            )
            return

        proceed = ui.uiUtils_show_text_report(
            TITLE + " Preview",
            preview_text,
            ok_text="Apply",
            cancel_text="Cancel",
            width=860,
            height=620,
        )
        if not proceed:
            return

        result_plans = _apply_updates(
            doc,
            floors,
            points,
            settings["tolerance_internal"],
        )
        ui.uiUtils_show_text_report(
            TITLE + " Results",
            _build_result_report(result_plans),
            ok_text="Close",
            cancel_text=None,
            width=860,
            height=620,
        )
    except Exception as exc:
        UI.TaskDialog.Show(TITLE + " - Error", "{}\n\n{}".format(exc, traceback.format_exc()))


if __name__ == "__main__":
    main()
