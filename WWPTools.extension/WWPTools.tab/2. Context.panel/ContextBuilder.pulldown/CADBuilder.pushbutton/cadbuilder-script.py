#! python3
import collections
from collections import defaultdict
import os
import random
import sys
import traceback

try:
    from collections.abc import Callable as _Callable
except Exception:
    _Callable = None

if _Callable is not None and not hasattr(collections, "Callable"):
    collections.Callable = _Callable
if not hasattr(collections, "callable"):
    collections.callable = callable

from System.Collections.Generic import List

import clr

from pyrevit import DB, revit, script

clr.AddReference("RevitAPIUI")
from Autodesk.Revit import UI


TITLE = "CAD Builder"
APP_ID = "WWPTools.CADBuilder"
DEFAULT_MIN_HEIGHT_M = 12.0
DEFAULT_MAX_HEIGHT_M = 60.0
MAX_PREVIEW_ITEMS = 20
MAX_FAILURE_ITEMS = 30


def _load_uiutils():
    script_dir = os.path.dirname(__file__)
    lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
    if lib_path not in sys.path:
        sys.path.append(lib_path)
    import WWP_uiUtils as ui
    return ui


def _meters_to_internal(value_m):
    try:
        return DB.UnitUtils.ConvertToInternalUnits(float(value_m), DB.UnitTypeId.Meters)
    except Exception:
        return DB.UnitUtils.ConvertToInternalUnits(float(value_m), DB.DisplayUnitType.DUT_METERS)


def _parse_number(raw_text):
    text = (raw_text or "").strip().replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        return None


def _show_alert(title, message, ui):
    ui.uiUtils_alert(message, title=title)


def _select_single(title, options, ui, prompt="Select an option:"):
    if not options:
        return None
    indices = ui.uiUtils_select_indices(
        options,
        title=title,
        prompt=prompt,
        multiselect=False,
        width=520,
        height=540,
    )
    if not indices:
        return None
    index = indices[0]
    if 0 <= index < len(options):
        return options[index]
    return None


def _select_multi(title, options, ui, prompt="Select items:"):
    if not options:
        return []
    indices = ui.uiUtils_select_indices(
        options,
        title=title,
        prompt=prompt,
        multiselect=True,
        width=520,
        height=560,
    )
    return [options[i] for i in indices if 0 <= i < len(options)]


def _prompt_number(title, prompt, default_value, ui):
    raw_value = ui.uiUtils_prompt_text(
        title=title,
        prompt=prompt,
        default_value=str(default_value),
        ok_text="OK",
        cancel_text="Cancel",
        width=420,
        height=220,
    )
    if raw_value is None:
        return None
    value = _parse_number(raw_value)
    if value is None:
        _show_alert(TITLE, "Invalid numeric value:\n{}".format(raw_value), ui)
        return None
    return value


def _pick_import_instance(uidoc):
    sel_ids = list(uidoc.Selection.GetElementIds())
    if len(sel_ids) == 1:
        elem = uidoc.Document.GetElement(sel_ids[0])
        if isinstance(elem, DB.ImportInstance):
            return elem

    class _ImportInstanceFilter(UI.Selection.ISelectionFilter):
        def AllowElement(self, element):
            return isinstance(element, DB.ImportInstance)

        def AllowReference(self, reference, position):
            return False

    try:
        ref = uidoc.Selection.PickObject(
            UI.Selection.ObjectType.Element,
            _ImportInstanceFilter(),
            "Pick a CAD import/link",
        )
    except Exception:
        return None

    return uidoc.Document.GetElement(ref.ElementId)


def _format_import_label(element):
    view_name = "<all views>"
    try:
        owner_view_id = element.OwnerViewId
        if owner_view_id and owner_view_id != DB.ElementId.InvalidElementId:
            owner_view = element.Document.GetElement(owner_view_id)
            if owner_view is not None:
                view_name = owner_view.Name
    except Exception:
        pass
    return "{} [Id: {} | View: {}]".format(element.Name, element.Id.IntegerValue, view_name)


def _choose_import_instance(doc, view, ui):
    picked = _pick_import_instance(revit.uidoc)
    if picked:
        return picked

    in_view = list(
        DB.FilteredElementCollector(doc, view.Id)
        .OfClass(DB.ImportInstance)
        .WhereElementIsNotElementType()
    )
    in_doc = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ImportInstance)
        .WhereElementIsNotElementType()
    )
    candidates = in_view or in_doc
    if not candidates:
        return None

    labels = []
    lookup = {}
    for element in sorted(candidates, key=lambda item: (item.Name or "", item.Id.IntegerValue)):
        label = _format_import_label(element)
        labels.append(label)
        lookup[label] = element

    selected = _select_single("Select CAD Import", labels, ui, prompt="Pick a CAD import/link:")
    if not selected:
        return None
    return lookup.get(selected)


def _get_layer_names(import_instance):
    cat = import_instance.Category
    if cat and cat.SubCategories:
        return sorted([subcat.Name for subcat in cat.SubCategories if subcat.Name])
    return []


def _walk_geometry(geom_elem, transform):
    for obj in geom_elem:
        if isinstance(obj, DB.GeometryInstance):
            next_transform = transform.Multiply(obj.Transform) if transform else obj.Transform
            symbol_geometry = obj.SymbolGeometry
            if symbol_geometry is None:
                continue
            for sub_obj in _walk_geometry(symbol_geometry, next_transform):
                yield sub_obj
        else:
            yield obj, transform


def _get_layer_name(doc, geom_obj):
    if geom_obj is None:
        return None
    try:
        graphics_style_id = geom_obj.GraphicsStyleId
    except Exception:
        return None
    if graphics_style_id == DB.ElementId.InvalidElementId:
        return None
    graphics_style = doc.GetElement(graphics_style_id)
    if isinstance(graphics_style, DB.GraphicsStyle) and graphics_style.GraphicsStyleCategory:
        return graphics_style.GraphicsStyleCategory.Name
    return None


def _distance(point_a, point_b):
    try:
        return point_a.DistanceTo(point_b)
    except Exception:
        return 0.0


def _point_key(point, tolerance):
    tol = tolerance or 1e-6
    return (
        int(round(point.X / tol)),
        int(round(point.Y / tol)),
        int(round(point.Z / tol)),
    )


def _point_key_xy(point, tolerance):
    tol = tolerance or 1e-6
    return (
        int(round(point.X / tol)),
        int(round(point.Y / tol)),
    )


def _clean_points(points, tolerance):
    cleaned = []
    for point in points or []:
        if point is None:
            continue
        if not cleaned or _distance(cleaned[-1], point) > tolerance:
            cleaned.append(point)
    if len(cleaned) > 1 and _distance(cleaned[0], cleaned[-1]) <= tolerance:
        cleaned[-1] = cleaned[0]
    return cleaned


def _close_points(points, tolerance):
    cleaned = _clean_points(points, tolerance)
    if len(cleaned) < 3:
        return cleaned
    if _distance(cleaned[0], cleaned[-1]) <= tolerance:
        cleaned[-1] = cleaned[0]
    else:
        cleaned.append(cleaned[0])
    return cleaned


def _polyline_points(polyline, transform, tolerance):
    points = list(polyline.GetCoordinates())
    if transform:
        points = [transform.OfPoint(point) for point in points]
    points = _clean_points(points, tolerance)
    is_closed = False
    try:
        is_closed = bool(polyline.IsClosed)
    except Exception:
        is_closed = len(points) > 2 and _distance(points[0], points[-1]) <= tolerance
    if is_closed and points and _distance(points[0], points[-1]) > tolerance:
        points.append(points[0])
    return points


def _curve_points(curve, transform, tolerance):
    try:
        curve = curve.CreateTransformed(transform) if transform else curve
    except Exception:
        pass

    try:
        points = list(curve.Tessellate())
    except Exception:
        points = []

    try:
        start_point = curve.GetEndPoint(0)
        end_point = curve.GetEndPoint(1)
        if not points:
            points = [start_point, end_point]
        else:
            if _distance(points[0], start_point) > tolerance:
                points.insert(0, start_point)
            if _distance(points[-1], end_point) > tolerance:
                points.append(end_point)
    except Exception:
        pass

    return _clean_points(points, tolerance)


def _collect_geometry(doc, import_instance, selected_layers, tolerance):
    options = DB.Options()
    options.IncludeNonVisibleObjects = True
    options.DetailLevel = DB.ViewDetailLevel.Fine

    geometry = import_instance.get_Geometry(options)
    if not geometry:
        return {}, {}

    rings_by_layer = defaultdict(list)
    segments_by_layer = defaultdict(list)
    stats_by_layer = defaultdict(lambda: {"objects": 0, "closed": 0, "segments": 0})

    for geom_obj, transform in _walk_geometry(geometry, None):
        if geom_obj is None:
            continue
        layer_name = _get_layer_name(doc, geom_obj)
        if not layer_name or layer_name not in selected_layers:
            continue

        stats = stats_by_layer[layer_name]
        stats["objects"] += 1

        points = None
        if isinstance(geom_obj, DB.PolyLine):
            points = _polyline_points(geom_obj, transform, tolerance)
        elif isinstance(geom_obj, DB.Curve):
            points = _curve_points(geom_obj, transform, tolerance)
        else:
            continue

        if len(points) < 2:
            continue

        is_closed = len(points) >= 3 and _distance(points[0], points[-1]) <= tolerance
        if is_closed:
            ring = _close_points(points, tolerance)
            if len(ring) >= 4:
                rings_by_layer[layer_name].append(ring)
                stats["closed"] += 1
            continue

        for index in range(len(points) - 1):
            start_point = points[index]
            end_point = points[index + 1]
            if _distance(start_point, end_point) <= tolerance:
                continue
            segments_by_layer[layer_name].append((start_point, end_point))
            stats["segments"] += 1

    return {"rings": rings_by_layer, "segments": segments_by_layer}, stats_by_layer


def _build_rings_from_segments(segments, tolerance):
    unique_segments = []
    seen_keys = set()
    for start_point, end_point in segments or []:
        if _distance(start_point, end_point) <= tolerance:
            continue
        start_key = _point_key(start_point, tolerance)
        end_key = _point_key(end_point, tolerance)
        if start_key == end_key:
            continue
        segment_key = tuple(sorted((start_key, end_key)))
        if segment_key in seen_keys:
            continue
        seen_keys.add(segment_key)
        unique_segments.append((start_point, end_point))

    adjacency = defaultdict(list)
    for index, (start_point, end_point) in enumerate(unique_segments):
        adjacency[_point_key(start_point, tolerance)].append((index, True))
        adjacency[_point_key(end_point, tolerance)].append((index, False))

    used = [False] * len(unique_segments)
    rings = []
    discarded = 0

    for index, (start_point, end_point) in enumerate(unique_segments):
        if used[index]:
            continue

        used[index] = True
        chain = [start_point, end_point]

        while True:
            if len(chain) >= 4 and _point_key(chain[-1], tolerance) == _point_key(chain[0], tolerance):
                ring = _close_points(chain, tolerance)
                if len(ring) >= 4:
                    rings.append(ring)
                else:
                    discarded += 1
                break

            current_key = _point_key(chain[-1], tolerance)
            previous_key = _point_key(chain[-2], tolerance) if len(chain) > 1 else None

            forward_choice = None
            fallback_choice = None
            for candidate_index, use_start in adjacency.get(current_key, []):
                if used[candidate_index]:
                    continue
                seg_start, seg_end = unique_segments[candidate_index]
                next_point = seg_end if use_start else seg_start
                next_key = _point_key(next_point, tolerance)
                candidate = (candidate_index, next_point)
                if previous_key is not None and next_key == previous_key:
                    if fallback_choice is None:
                        fallback_choice = candidate
                    continue
                forward_choice = candidate
                break

            chosen = forward_choice or fallback_choice
            if chosen is None:
                discarded += 1
                break

            candidate_index, next_point = chosen
            used[candidate_index] = True
            chain.append(next_point)

    return rings, discarded


def _signed_area_xy(points):
    area = 0.0
    for index in range(len(points) - 1):
        point_a = points[index]
        point_b = points[index + 1]
        area += (point_a.X * point_b.Y) - (point_b.X * point_a.Y)
    return area / 2.0


def _point_in_ring_xy(point, ring):
    x_value = point.X
    y_value = point.Y
    inside = False
    for index in range(len(ring) - 1):
        point_a = ring[index]
        point_b = ring[index + 1]
        intersects = (point_a.Y > y_value) != (point_b.Y > y_value)
        if not intersects:
            continue
        slope_x = ((point_b.X - point_a.X) * (y_value - point_a.Y) / ((point_b.Y - point_a.Y) or 1e-12)) + point_a.X
        if x_value < slope_x:
            inside = not inside
    return inside


def _ring_average_elevation(ring):
    points = ring[:-1] if len(ring) > 1 else ring
    if not points:
        return 0.0
    return sum(point.Z for point in points) / float(len(points))


def _ring_is_horizontal(ring, tolerance):
    points = ring[:-1] if len(ring) > 1 else ring
    if not points:
        return False
    base_z = _ring_average_elevation(ring)
    for point in points:
        if abs(point.Z - base_z) > tolerance:
            return False
    return True


def _normalize_ring(ring, tolerance, min_area_internal):
    closed = _close_points(ring, tolerance)
    unique_points = {_point_key_xy(point, tolerance) for point in closed[:-1]} if len(closed) > 1 else set()
    if len(closed) < 4 or len(unique_points) < 3:
        return None
    if abs(_signed_area_xy(closed)) < min_area_internal:
        return None
    return closed


def _rings_to_features(layer_name, rings, tolerance, min_area_internal, feature_prefix):
    records = []
    for ring_index, ring in enumerate(rings or []):
        normalized = _normalize_ring(ring, tolerance, min_area_internal)
        if normalized is None:
            continue
        area = abs(_signed_area_xy(normalized))
        records.append(
            {
                "id": "{}:{}".format(feature_prefix, ring_index + 1),
                "layer": layer_name,
                "points": normalized,
                "area": area,
                "parent": None,
                "depth": 0,
                "children": [],
            }
        )

    order = sorted(range(len(records)), key=lambda idx: records[idx]["area"], reverse=True)
    for idx in order:
        point = records[idx]["points"][0]
        parent_index = None
        for candidate_idx in order:
            if records[candidate_idx]["area"] <= records[idx]["area"]:
                continue
            if _point_in_ring_xy(point, records[candidate_idx]["points"]):
                if parent_index is None or records[candidate_idx]["area"] < records[parent_index]["area"]:
                    parent_index = candidate_idx
        records[idx]["parent"] = parent_index
        if parent_index is not None:
            records[idx]["depth"] = records[parent_index]["depth"] + 1
            records[parent_index]["children"].append(idx)

    features = []
    for idx, record in enumerate(records):
        if record["depth"] % 2 != 0:
            continue
        features.append(
            {
                "id": record["id"],
                "layer": layer_name,
                "outer": record["points"],
                "inners": [records[child_idx]["points"] for child_idx in record["children"] if records[child_idx]["depth"] == record["depth"] + 1],
                "base_elevation": _ring_average_elevation(record["points"]),
            }
        )

    return features


def _orient_ring(points, clockwise):
    area = _signed_area_xy(points)
    is_clockwise = area < 0
    if is_clockwise != clockwise:
        points = list(reversed(points))
    return points


def _ring_to_curve_loop(points, clockwise, base_elevation, tolerance):
    ring = _close_points(points, tolerance)
    unique_points = {_point_key_xy(point, tolerance) for point in ring[:-1]} if len(ring) > 1 else set()
    if len(ring) < 4 or len(unique_points) < 3:
        return None

    ordered_points = _orient_ring(ring, clockwise=clockwise)
    curve_loop = DB.CurveLoop()
    for index in range(len(ordered_points) - 1):
        start_point = ordered_points[index]
        end_point = ordered_points[index + 1]
        start_xyz = DB.XYZ(start_point.X, start_point.Y, base_elevation)
        end_xyz = DB.XYZ(end_point.X, end_point.Y, base_elevation)
        if start_xyz.IsAlmostEqualTo(end_xyz):
            continue
        curve_loop.Append(DB.Line.CreateBound(start_xyz, end_xyz))

    return curve_loop if curve_loop.GetExactLength() > 0 else None


def _feature_to_curve_loops(feature, tolerance):
    base_elevation = feature.get("base_elevation") or 0.0
    outer_loop = _ring_to_curve_loop(feature.get("outer") or [], clockwise=False, base_elevation=base_elevation, tolerance=tolerance)
    if outer_loop is None:
        return None

    curve_loops = List[DB.CurveLoop]()
    curve_loops.Add(outer_loop)

    for inner_ring in feature.get("inners") or []:
        inner_loop = _ring_to_curve_loop(inner_ring, clockwise=True, base_elevation=base_elevation, tolerance=tolerance)
        if inner_loop is not None:
            curve_loops.Add(inner_loop)

    return curve_loops


def _set_comment(element, text):
    try:
        parameter = element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if parameter and not parameter.IsReadOnly:
            parameter.Set(text)
            return
    except Exception:
        pass

    try:
        parameter = element.LookupParameter("Comments")
        if parameter and not parameter.IsReadOnly:
            parameter.Set(text)
    except Exception:
        pass


def _create_solid(curve_loops, height_m):
    height_internal = _meters_to_internal(height_m)
    if height_internal <= 0:
        raise Exception("Computed height is not positive.")
    return DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        curve_loops,
        DB.XYZ.BasisZ,
        height_internal,
    )


def _create_mass(doc, feature, import_instance_name, tolerance):
    outer_ring = feature.get("outer") or []
    if not outer_ring:
        raise Exception("Missing footprint geometry.")

    if not _ring_is_horizontal(outer_ring, _meters_to_internal(0.05)):
        raise Exception("Footprint is not horizontal.")

    curve_loops = _feature_to_curve_loops(feature, tolerance)
    if curve_loops is None or curve_loops.Count == 0:
        raise Exception("Failed to build a valid footprint loop.")

    solid = _create_solid(curve_loops, feature["height_m"])
    errors = []
    for built_in_category in (DB.BuiltInCategory.OST_Mass, DB.BuiltInCategory.OST_GenericModel):
        shape = None
        try:
            shape = DB.DirectShape.CreateElement(doc, DB.ElementId(built_in_category))
            shape.ApplicationId = APP_ID
            shape.ApplicationDataId = feature["id"]
            geometry = List[DB.GeometryObject]()
            geometry.Add(solid)
            shape.SetShape(geometry)
            _set_comment(
                shape,
                "CAD Builder | {} | Layer: {} | Height: {:.2f} m".format(
                    import_instance_name,
                    feature["layer"],
                    feature["height_m"],
                ),
            )
            return str(built_in_category)
        except Exception as ex:
            errors.append("{}: {}".format(str(built_in_category), str(ex)))
            if shape is not None:
                try:
                    doc.Delete(shape.Id)
                except Exception:
                    pass

    raise Exception("; ".join(errors[:4]))


def _build_preview(import_instance, selected_layers, features, min_height_m, max_height_m, layer_feature_counts, discarded_counts):
    lines = [
        "CAD import: {}".format(import_instance.Name),
        "Selected layers: {}".format(", ".join(selected_layers)),
        "Detected footprints: {}".format(len(features)),
        "Random height range (m): {:.2f} - {:.2f}".format(min_height_m, max_height_m),
        "",
        "By layer:",
    ]

    for layer_name in selected_layers:
        lines.append(
            "- {}: {} footprints".format(
                layer_name,
                layer_feature_counts.get(layer_name, 0),
            )
        )

    discarded_total = sum(discarded_counts.values())
    if discarded_total:
        lines.extend(
            [
                "",
                "Discarded open/non-closed fragments:",
            ]
        )
        for layer_name in selected_layers:
            count = discarded_counts.get(layer_name, 0)
            if count:
                lines.append("- {}: {}".format(layer_name, count))

    lines.extend(["", "Preview of first {} masses:".format(MAX_PREVIEW_ITEMS)])
    for feature in features[:MAX_PREVIEW_ITEMS]:
        lines.append(
            "- {} | {} | {:.2f} m".format(
                feature["layer"],
                feature["id"],
                feature["height_m"],
            )
        )

    if len(features) > MAX_PREVIEW_ITEMS:
        lines.append("- ...")

    return "\n".join(lines)


def _build_result_text(created_total, category_counts, layer_created_counts, failures):
    lines = [
        "Created masses: {}".format(created_total),
        "Mass category: {}".format(category_counts.get("OST_Mass", 0)),
        "Generic Model fallback: {}".format(category_counts.get("OST_GenericModel", 0)),
        "",
        "Created by layer:",
    ]

    for layer_name in sorted(layer_created_counts.keys()):
        lines.append("- {}: {}".format(layer_name, layer_created_counts[layer_name]))

    lines.append("")
    lines.append("Failures: {}".format(len(failures)))
    if failures:
        lines.append("")
        lines.append("First {} failures:".format(MAX_FAILURE_ITEMS))
        for failure in failures[:MAX_FAILURE_ITEMS]:
            lines.append("- {} | {}".format(failure["id"], failure["reason"]))

    return "\n".join(lines)


def _report_to_output(text):
    try:
        output = script.get_output()
        for line in text.splitlines():
            output.print_md(line)
    except Exception:
        print(text)


def main():
    doc = revit.doc
    if doc is None:
        UI.TaskDialog.Show(TITLE, "No active Revit document found.")
        return

    ui = _load_uiutils()
    import_instance = _choose_import_instance(doc, doc.ActiveView, ui)
    if import_instance is None:
        _show_alert(TITLE, "No CAD import/link selected.", ui)
        return

    layer_names = _get_layer_names(import_instance)
    if not layer_names:
        _show_alert(TITLE, "No layers found on the selected CAD import.", ui)
        return

    selected_layers = _select_multi("Select CAD Layers", layer_names, ui, prompt="Choose CAD layers to convert into massing:")
    if not selected_layers:
        return

    min_height_m = _prompt_number(TITLE, "Minimum random height (m):", DEFAULT_MIN_HEIGHT_M, ui)
    if min_height_m is None:
        return

    max_height_m = _prompt_number(TITLE, "Maximum random height (m):", DEFAULT_MAX_HEIGHT_M, ui)
    if max_height_m is None:
        return

    if min_height_m <= 0 or max_height_m <= 0:
        _show_alert(TITLE, "Random heights must be greater than zero.", ui)
        return

    if max_height_m < min_height_m:
        min_height_m, max_height_m = max_height_m, min_height_m

    point_tolerance = max(doc.Application.ShortCurveTolerance, _meters_to_internal(0.01))
    min_area_internal = _meters_to_internal(1.0) * _meters_to_internal(1.0)
    geometry_by_layer, stats_by_layer = _collect_geometry(doc, import_instance, set(selected_layers), point_tolerance)

    features = []
    discarded_counts = {}
    layer_feature_counts = {}
    for layer_name in selected_layers:
        direct_rings = list((geometry_by_layer.get("rings") or {}).get(layer_name, []))
        segment_rings, discarded_count = _build_rings_from_segments(
            (geometry_by_layer.get("segments") or {}).get(layer_name, []),
            point_tolerance,
        )
        layer_features = _rings_to_features(
            layer_name,
            direct_rings + segment_rings,
            point_tolerance,
            min_area_internal,
            "cad:{}:{}".format(import_instance.Id.IntegerValue, layer_name),
        )
        features.extend(layer_features)
        discarded_counts[layer_name] = discarded_count
        layer_feature_counts[layer_name] = len(layer_features)

    if not features:
        lines = ["No closed footprints were found on the selected layers."]
        for layer_name in selected_layers:
            stats = stats_by_layer.get(layer_name) or {}
            lines.append(
                "- {} | closed objects: {} | open segments: {}".format(
                    layer_name,
                    stats.get("closed", 0),
                    stats.get("segments", 0),
                )
            )
        _show_alert(TITLE, "\n".join(lines), ui)
        return

    rng = random.Random()
    for feature in features:
        feature["height_m"] = rng.uniform(min_height_m, max_height_m)

    preview_text = _build_preview(
        import_instance,
        selected_layers,
        features,
        min_height_m,
        max_height_m,
        layer_feature_counts,
        discarded_counts,
    )
    if not ui.uiUtils_show_text_report(
        "{} - Preview".format(TITLE),
        preview_text,
        ok_text="Build",
        cancel_text="Cancel",
        width=760,
        height=620,
    ):
        return

    created_total = 0
    category_counts = {"OST_Mass": 0, "OST_GenericModel": 0}
    layer_created_counts = defaultdict(int)
    failures = []

    transaction = DB.Transaction(doc, TITLE)
    started = False
    try:
        transaction.Start()
        started = True

        for feature in features:
            try:
                category_name = _create_mass(doc, feature, import_instance.Name, point_tolerance)
                category_counts[category_name] = category_counts.get(category_name, 0) + 1
                layer_created_counts[feature["layer"]] += 1
                created_total += 1
            except Exception as ex:
                failures.append({"id": feature["id"], "reason": str(ex)})

        transaction.Commit()
    except Exception as ex:
        if started:
            try:
                transaction.RollBack()
            except Exception:
                pass
        UI.TaskDialog.Show(TITLE + " - Error", "{}\n\n{}".format(ex, traceback.format_exc()))
        return

    result_text = _build_result_text(created_total, category_counts, layer_created_counts, failures)
    _report_to_output(result_text)
    ui.uiUtils_show_text_report(
        "{} - Results".format(TITLE),
        result_text,
        ok_text="Close",
        cancel_text=None,
        width=760,
        height=560,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        UI.TaskDialog.Show(TITLE + " - Error", "{}\n\n{}".format(exc, traceback.format_exc()))
