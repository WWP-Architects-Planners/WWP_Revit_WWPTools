import clr
import math
import os
import random
import re
import traceback
import xml.etree.ElementTree as ET

from System.Collections.Generic import List

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit import DB, UI

import WWP_uiUtils as ui


TITLE = "Building Importer"
APP_ID = "WWPTools.BuildingImporter"
DEFAULT_MIN_HEIGHT_M = 4.0
DEFAULT_MAX_HEIGHT_M = 6.0
LEVEL_HEIGHT_INCREMENT_M = 2.95
GROUND_LEVEL_HEIGHT_M = 3.25
MAX_FAILURES_IN_REPORT = 25


def _get_doc():
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is None:
            return None
        return uidoc.Document
    except Exception:
        return None


def _strip_ns(tag_name):
    if "}" in tag_name:
        return tag_name.split("}", 1)[1]
    return tag_name


def _meters_to_internal(value_m):
    try:
        return DB.UnitUtils.ConvertToInternalUnits(float(value_m), DB.UnitTypeId.Meters)
    except Exception:
        return DB.UnitUtils.ConvertToInternalUnits(float(value_m), DB.DisplayUnitType.DUT_METERS)


def _normalize_number_text(raw_text):
    text = (raw_text or "").strip()
    if not text:
        return text
    if "," in text and "." not in text and text.count(",") == 1:
        return text.replace(",", ".")
    return text


def _parse_first_number(raw_text):
    text = _normalize_number_text(raw_text)
    match = re.search(r"[-+]?\d*\.?\d+", text or "")
    if not match:
        return None
    try:
        return float(match.group(0))
    except Exception:
        return None


def _parse_height_m(tags):
    for key in ("height", "building:height", "est_height"):
        raw_value = tags.get(key)
        if not raw_value:
            continue

        value = _parse_first_number(raw_value)
        if value is None or value <= 0:
            continue

        text = raw_value.lower()
        if "ft" in text or "feet" in text or "'" in text:
            return value * 0.3048
        if "mm" in text:
            return value / 1000.0
        if "cm" in text:
            return value / 100.0
        return value
    return None


def _parse_levels(tags):
    for key in ("building:levels", "levels"):
        raw_value = tags.get(key)
        if not raw_value:
            continue
        value = _parse_first_number(raw_value)
        if value is None or value <= 0:
            continue
        return value
    return None


def _fallback_height_from_levels(levels):
    if levels is None or levels <= 0:
        return None
    return ((levels - 1.0) * LEVEL_HEIGHT_INCREMENT_M) + GROUND_LEVEL_HEIGHT_M


def _resolve_height_m(tags, rng, min_height_m, max_height_m):
    explicit_height_m = _parse_height_m(tags)
    if explicit_height_m is not None:
        return explicit_height_m, "height"

    levels = _parse_levels(tags)
    level_height_m = _fallback_height_from_levels(levels)
    if level_height_m is not None:
        return level_height_m, "levels"

    return rng.uniform(min_height_m, max_height_m), "random"


def _is_building(tags):
    for key in ("building", "building:part"):
        value = (tags.get(key) or "").strip().lower()
        if value and value not in ("no", "false", "0"):
            return True
    return False


def _prompt_number(prompt, default_value):
    raw_value = ui.uiUtils_prompt_text(
        title=TITLE,
        prompt=prompt,
        default_value=str(default_value),
        ok_text="OK",
        cancel_text="Cancel",
        width=420,
        height=220,
    )
    if raw_value is None:
        return None

    value = _parse_first_number(raw_value)
    if value is None:
        UI.TaskDialog.Show(TITLE, "Invalid numeric value for:\n{}".format(prompt))
        return None
    return value


def _prompt_inputs(doc):
    file_path = ui.uiUtils_open_file_dialog(
        title="Select OSM File",
        filter_text="OpenStreetMap (*.osm;*.xml)|*.osm;*.xml|All files (*.*)|*.*",
        multiselect=False,
    )
    if not file_path:
        return None

    min_height_m = _prompt_number("Minimum fallback height (m):", DEFAULT_MIN_HEIGHT_M)
    if min_height_m is None:
        return None

    max_height_m = _prompt_number("Maximum fallback height (m):", DEFAULT_MAX_HEIGHT_M)
    if max_height_m is None:
        return None

    material_id = _prompt_material(doc)
    if material_id is None:
        return None

    if max_height_m < min_height_m:
        min_height_m, max_height_m = max_height_m, min_height_m

    return {
        "file_path": file_path,
        "min_height_m": min_height_m,
        "max_height_m": max_height_m,
        "material_id": material_id,
    }


def _prompt_material(doc):
    materials = list(DB.FilteredElementCollector(doc).OfClass(DB.Material))
    materials.sort(key=lambda material: material.Name.lower())

    names = ["<By Category>"]
    ids = [DB.ElementId.InvalidElementId]
    for material in materials:
        names.append(material.Name)
        ids.append(material.Id)

    selected_indices = ui.uiUtils_select_indices(
        names,
        title=TITLE,
        prompt="Select material for imported buildings:",
        multiselect=False,
        width=520,
        height=640,
    )
    if not selected_indices:
        return None
    return ids[selected_indices[0]]


def _load_osm(path):
    tree = ET.parse(path)
    root = tree.getroot()

    nodes = {}
    ways = {}
    relations = []
    bounds = None

    for element in root:
        element_name = _strip_ns(element.tag)

        if element_name == "bounds":
            bounds = {
                "minlat": float(element.attrib.get("minlat", 0.0)),
                "minlon": float(element.attrib.get("minlon", 0.0)),
                "maxlat": float(element.attrib.get("maxlat", 0.0)),
                "maxlon": float(element.attrib.get("maxlon", 0.0)),
            }
            continue

        if element_name == "node":
            node_id = element.attrib.get("id")
            if not node_id:
                continue
            nodes[node_id] = (
                float(element.attrib.get("lat", 0.0)),
                float(element.attrib.get("lon", 0.0)),
            )
            continue

        if element_name == "way":
            way_id = element.attrib.get("id")
            refs = []
            tags = {}
            for child in element:
                child_name = _strip_ns(child.tag)
                if child_name == "nd":
                    refs.append(child.attrib.get("ref"))
                elif child_name == "tag":
                    tags[child.attrib.get("k")] = child.attrib.get("v")
            if way_id:
                ways[way_id] = {"refs": refs, "tags": tags}
            continue

        if element_name == "relation":
            relation_id = element.attrib.get("id")
            members = []
            tags = {}
            for child in element:
                child_name = _strip_ns(child.tag)
                if child_name == "member":
                    members.append(
                        {
                            "type": child.attrib.get("type"),
                            "ref": child.attrib.get("ref"),
                            "role": (child.attrib.get("role") or "").strip().lower(),
                        }
                    )
                elif child_name == "tag":
                    tags[child.attrib.get("k")] = child.attrib.get("v")
            if relation_id:
                relations.append({"id": relation_id, "members": members, "tags": tags})

    return nodes, ways, relations, bounds


def _get_projection_origin(bounds, nodes):
    if bounds:
        return (
            (bounds["minlat"] + bounds["maxlat"]) / 2.0,
            (bounds["minlon"] + bounds["maxlon"]) / 2.0,
        )

    if not nodes:
        return 0.0, 0.0

    total_lat = 0.0
    total_lon = 0.0
    count = 0
    for lat, lon in nodes.values():
        total_lat += lat
        total_lon += lon
        count += 1
    return total_lat / float(count), total_lon / float(count)


def _project_latlon(lat, lon, origin_lat, origin_lon):
    radius_m = 6378137.0
    lat_rad = math.radians(lat)
    lon_rad = math.radians(lon)
    origin_lat_rad = math.radians(origin_lat)
    origin_lon_rad = math.radians(origin_lon)
    x_m = (lon_rad - origin_lon_rad) * math.cos(origin_lat_rad) * radius_m
    y_m = (lat_rad - origin_lat_rad) * radius_m
    return x_m, y_m


def _remove_consecutive_duplicate_points(points, tolerance_m=0.01):
    cleaned = []
    for point in points:
        if not cleaned:
            cleaned.append(point)
            continue

        prev_x, prev_y = cleaned[-1]
        curr_x, curr_y = point
        if math.hypot(curr_x - prev_x, curr_y - prev_y) > tolerance_m:
            cleaned.append(point)

    if len(cleaned) > 1:
        first_x, first_y = cleaned[0]
        last_x, last_y = cleaned[-1]
        if math.hypot(first_x - last_x, first_y - last_y) <= tolerance_m:
            cleaned[-1] = cleaned[0]

    return cleaned


def _refs_to_projected_ring(refs, nodes, origin_lat, origin_lon):
    points = []
    for node_ref in refs:
        lat_lon = nodes.get(node_ref)
        if lat_lon is None:
            return None
        lat, lon = lat_lon
        points.append(_project_latlon(lat, lon, origin_lat, origin_lon))

    points = _remove_consecutive_duplicate_points(points)
    if len(points) < 3:
        return None

    if points[0] != points[-1]:
        points.append(points[0])

    unique_points = set(points[:-1])
    if len(unique_points) < 3:
        return None

    return points


def _assemble_rings(way_ids, ways):
    segments = []
    for way_id in way_ids:
        way = ways.get(way_id)
        if not way:
            continue

        refs = [node_ref for node_ref in (way.get("refs") or []) if node_ref]
        if len(refs) < 2:
            continue
        segments.append(list(refs))

    rings = []
    while segments:
        chain = segments.pop(0)
        progressed = True
        while progressed and chain and chain[0] != chain[-1]:
            progressed = False
            for index, refs in enumerate(segments):
                start_ref = refs[0]
                end_ref = refs[-1]

                if chain[-1] == start_ref:
                    chain.extend(refs[1:])
                elif chain[-1] == end_ref:
                    chain.extend(list(reversed(refs[:-1])))
                elif chain[0] == end_ref:
                    chain = refs[:-1] + chain
                elif chain[0] == start_ref:
                    chain = list(reversed(refs[1:])) + chain
                else:
                    continue

                segments.pop(index)
                progressed = True
                break

        if chain and chain[0] == chain[-1] and len(set(chain[:-1])) >= 3:
            rings.append(chain)

    return rings


def _signed_area(points):
    area = 0.0
    for index in range(len(points) - 1):
        x1, y1 = points[index]
        x2, y2 = points[index + 1]
        area += (x1 * y2) - (x2 * y1)
    return area / 2.0


def _point_in_polygon(point, polygon):
    x, y = point
    inside = False
    for index in range(len(polygon) - 1):
        x1, y1 = polygon[index]
        x2, y2 = polygon[index + 1]
        intersects = ((y1 > y) != (y2 > y))
        if not intersects:
            continue
        slope_x = ((x2 - x1) * (y - y1) / ((y2 - y1) or 1e-12)) + x1
        if x < slope_x:
            inside = not inside
    return inside


def _build_features(nodes, ways, relations, bounds):
    origin_lat, origin_lon = _get_projection_origin(bounds, nodes)
    features = []
    relation_way_ids = set()

    for relation in relations:
        tags = relation.get("tags") or {}
        if not _is_building(tags):
            continue

        outer_way_ids = []
        inner_way_ids = []
        for member in relation.get("members") or []:
            if member.get("type") != "way":
                continue

            role = member.get("role") or ""
            if role == "inner":
                inner_way_ids.append(member.get("ref"))
            else:
                outer_way_ids.append(member.get("ref"))

        outer_rings = _assemble_rings(outer_way_ids, ways)
        if not outer_rings:
            continue

        inner_rings = _assemble_rings(inner_way_ids, ways)
        for way_id in outer_way_ids + inner_way_ids:
            if way_id:
                relation_way_ids.add(way_id)

        projected_outers = []
        for ring in outer_rings:
            projected = _refs_to_projected_ring(ring, nodes, origin_lat, origin_lon)
            if projected:
                projected_outers.append(projected)

        projected_inners = []
        for ring in inner_rings:
            projected = _refs_to_projected_ring(ring, nodes, origin_lat, origin_lon)
            if projected:
                projected_inners.append(projected)

        for outer_index, outer_ring in enumerate(projected_outers):
            assigned_inners = []
            for inner_ring in projected_inners:
                if _point_in_polygon(inner_ring[0], outer_ring):
                    assigned_inners.append(inner_ring)

            features.append(
                {
                    "id": "relation/{}:{}".format(relation["id"], outer_index + 1),
                    "tags": tags,
                    "outer": outer_ring,
                    "inners": assigned_inners,
                }
            )

    for way_id, way in ways.items():
        if way_id in relation_way_ids:
            continue

        tags = way.get("tags") or {}
        if not _is_building(tags):
            continue

        refs = [node_ref for node_ref in (way.get("refs") or []) if node_ref]
        if len(refs) < 4 or refs[0] != refs[-1]:
            continue

        projected = _refs_to_projected_ring(refs, nodes, origin_lat, origin_lon)
        if not projected:
            continue

        features.append(
            {
                "id": "way/{}".format(way_id),
                "tags": tags,
                "outer": projected,
                "inners": [],
            }
        )

    return features


def _orient_ring(points, clockwise):
    area = _signed_area(points)
    is_clockwise = area < 0
    if is_clockwise != clockwise:
        points = list(reversed(points))
    return points


def _ring_to_curve_loop(points, clockwise):
    ordered_points = _orient_ring(points, clockwise=clockwise)
    if not ordered_points or len(ordered_points) < 4:
        return None

    loop = DB.CurveLoop()
    open_points = ordered_points[:-1]
    for index, point in enumerate(open_points):
        next_point = open_points[(index + 1) % len(open_points)]
        start_xyz = DB.XYZ(_meters_to_internal(point[0]), _meters_to_internal(point[1]), 0.0)
        end_xyz = DB.XYZ(_meters_to_internal(next_point[0]), _meters_to_internal(next_point[1]), 0.0)
        if start_xyz.IsAlmostEqualTo(end_xyz):
            continue
        loop.Append(DB.Line.CreateBound(start_xyz, end_xyz))

    return loop if loop.GetExactLength() > 0 else None


def _feature_to_curve_loops(feature):
    curve_loops = List[DB.CurveLoop]()

    outer_loop = _ring_to_curve_loop(feature.get("outer") or [], clockwise=False)
    if outer_loop is None:
        return None
    curve_loops.Add(outer_loop)

    for inner_ring in feature.get("inners") or []:
        inner_loop = _ring_to_curve_loop(inner_ring, clockwise=True)
        if inner_loop is not None:
            curve_loops.Add(inner_loop)

    return curve_loops


def _set_material_parameter(element, material_id):
    if material_id is None or material_id == DB.ElementId.InvalidElementId:
        return False

    for built_in_param in (
        DB.BuiltInParameter.MATERIAL_ID_PARAM,
        DB.BuiltInParameter.STRUCTURAL_MATERIAL_PARAM,
    ):
        try:
            parameter = element.get_Parameter(built_in_param)
            if parameter and not parameter.IsReadOnly and parameter.StorageType == DB.StorageType.ElementId:
                parameter.Set(material_id)
                return True
        except Exception:
            pass

    for param_name in ("Material", "Structural Material"):
        try:
            parameter = element.LookupParameter(param_name)
            if parameter and not parameter.IsReadOnly and parameter.StorageType == DB.StorageType.ElementId:
                parameter.Set(material_id)
                return True
        except Exception:
            pass

    return False


def _set_comment(element, text):
    try:
        parameter = element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if parameter and not parameter.IsReadOnly:
            parameter.Set(text)
    except Exception:
        pass


def _create_solid(curve_loops, height_m, material_id):
    height_internal = _meters_to_internal(height_m)
    if height_internal <= 0:
        raise Exception("Computed height is not positive.")

    if material_id is not None and material_id != DB.ElementId.InvalidElementId:
        try:
            solid_options = DB.SolidOptions(material_id, DB.ElementId.InvalidElementId)
            return DB.GeometryCreationUtilities.CreateExtrusionGeometry(
                curve_loops,
                DB.XYZ.BasisZ,
                height_internal,
                solid_options,
            )
        except Exception:
            pass

    return DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        curve_loops,
        DB.XYZ.BasisZ,
        height_internal,
    )


def _create_direct_shape(doc, solid, feature_id, height_m, material_id):
    errors = []
    for built_in_category in (DB.BuiltInCategory.OST_Mass, DB.BuiltInCategory.OST_GenericModel):
        shape = None
        try:
            shape = DB.DirectShape.CreateElement(doc, DB.ElementId(built_in_category))
            shape.ApplicationId = APP_ID
            shape.ApplicationDataId = feature_id

            geometry = List[DB.GeometryObject]()
            geometry.Add(solid)
            shape.SetShape(geometry)

            _set_material_parameter(shape, material_id)
            _set_comment(shape, "OSM feature: {} | Height (m): {:.2f}".format(feature_id, height_m))
            return shape, built_in_category
        except Exception as ex:
            errors.append("{}: {}".format(str(built_in_category), str(ex)))
            if shape is not None:
                try:
                    doc.Delete(shape.Id)
                except Exception:
                    pass

    raise Exception("; ".join(errors))


def _format_preview(material_id, features, stats, min_height_m, max_height_m):
    material_name = "<By Category>"
    if material_id is not None and material_id != DB.ElementId.InvalidElementId:
        try:
            material_name = doc.GetElement(material_id).Name
        except Exception:
            pass

    lines = [
        "Buildings found: {}".format(len(features)),
        "Explicit height tags: {}".format(stats["height"]),
        "Derived from levels: {}".format(stats["levels"]),
        "Random fallback heights: {}".format(stats["random"]),
        "Fallback range (m): {:.2f} - {:.2f}".format(min_height_m, max_height_m),
        "Material: {}".format(material_name),
        "",
        "Preview of first 20 buildings:",
    ]

    for feature in features[:20]:
        lines.append(
            "{} | {:.2f} m | {}".format(
                feature["id"],
                feature["height_m"],
                feature["height_source"],
            )
        )

    if len(features) > 20:
        lines.append("...")

    return "\n".join(lines)


def _summarize_results(created_count, category_counts, failures):
    lines = [
        "Created: {}".format(created_count),
        "Mass category: {}".format(category_counts.get("OST_Mass", 0)),
        "Generic Model fallback: {}".format(category_counts.get("OST_GenericModel", 0)),
        "Failures: {}".format(len(failures)),
    ]

    if failures:
        lines.append("")
        lines.append("Failures (first {}):".format(MAX_FAILURES_IN_REPORT))
        for failure in failures[:MAX_FAILURES_IN_REPORT]:
            lines.append("{} | {}".format(failure["id"], failure["reason"]))

    return "\n".join(lines)


def main():
    global doc

    doc = _get_doc()
    if doc is None:
        UI.TaskDialog.Show(TITLE, "No active Revit document found.")
        return

    user_inputs = _prompt_inputs(doc)
    if user_inputs is None:
        return

    file_path = user_inputs["file_path"]
    if not os.path.isfile(file_path):
        UI.TaskDialog.Show(TITLE, "Selected file does not exist:\n{}".format(file_path))
        return

    nodes, ways, relations, bounds = _load_osm(file_path)
    features = _build_features(nodes, ways, relations, bounds)
    if not features:
        UI.TaskDialog.Show(TITLE, "No closed OSM building footprints were found.")
        return

    rng_seed = int(round(user_inputs["min_height_m"] * 1000.0))
    rng = random.Random(rng_seed)

    stats = {"height": 0, "levels": 0, "random": 0}
    for feature in features:
        height_m, height_source = _resolve_height_m(
            feature["tags"],
            rng,
            user_inputs["min_height_m"],
            user_inputs["max_height_m"],
        )
        feature["height_m"] = height_m
        feature["height_source"] = height_source
        stats[height_source] += 1

    preview_text = _format_preview(
        user_inputs["material_id"],
        features,
        stats,
        user_inputs["min_height_m"],
        user_inputs["max_height_m"],
    )
    if not ui.uiUtils_show_text_report(
        "{} - Preview".format(TITLE),
        preview_text,
        ok_text="Import",
        cancel_text="Cancel",
        width=760,
        height=560,
    ):
        return

    created_count = 0
    category_counts = {"OST_Mass": 0, "OST_GenericModel": 0}
    failures = []

    transaction = DB.Transaction(doc, TITLE)
    started = False
    try:
        transaction.Start()
        started = True

        for feature in features:
            try:
                curve_loops = _feature_to_curve_loops(feature)
                if curve_loops is None or curve_loops.Count == 0:
                    raise Exception("Failed to build a valid footprint loop.")

                solid = _create_solid(
                    curve_loops,
                    feature["height_m"],
                    user_inputs["material_id"],
                )
                _, built_in_category = _create_direct_shape(
                    doc,
                    solid,
                    feature["id"],
                    feature["height_m"],
                    user_inputs["material_id"],
                )

                created_count += 1
                category_counts[str(built_in_category)] = category_counts.get(str(built_in_category), 0) + 1
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

    result_text = _summarize_results(created_count, category_counts, failures)
    ui.uiUtils_show_text_report(
        "{} - Results".format(TITLE),
        result_text,
        ok_text="Close",
        cancel_text=None,
        width=760,
        height=520,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        UI.TaskDialog.Show(TITLE + " - Error", "{}\n\n{}".format(exc, traceback.format_exc()))
