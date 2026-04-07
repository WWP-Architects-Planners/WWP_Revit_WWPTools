# pyright: reportGeneralTypeIssues=false
# Combined tool: fixes off-axis sketch/model lines (Revit warnings) AND
# grid line angles AND grid spacing from dimension chains.

import math
import traceback

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit import DB

TITLE = "Fix All Angles"
ANGLE_PRECISION = 2
ANGLE_EPSILON = 1e-6
PERP_CANDIDATE_TOLERANCE_DEG = 0.2
CARDINAL_SNAP_TOLERANCE_DEG = 0.01
SHAKE_MM = 500.0

# ---------------------------------------------------------------------------
# Angle helpers
# ---------------------------------------------------------------------------

def _normalize_180(angle_degrees):
    value = angle_degrees % 180.0
    if value < 0.0:
        value += 180.0
    if abs(value - 180.0) <= 1e-16:
        return 0.0
    return value


def _shortest_delta_180(current_degrees, target_degrees):
    current = _normalize_180(current_degrees)
    target = _normalize_180(target_degrees)
    delta = target - current
    while delta <= -90.0:
        delta += 180.0
    while delta > 90.0:
        delta -= 180.0
    return delta


def _line_world_angle_180(line):
    if line is None:
        return None
    direction = line.Direction
    if abs(direction.X) < 1e-9 and abs(direction.Y) < 1e-9:
        return None
    angle = math.degrees(math.atan2(direction.Y, direction.X))
    return _normalize_180(angle)


def _snap_cardinal_angle(angle_degrees):
    angle = _normalize_180(angle_degrees)
    if abs(angle - 0.0) <= CARDINAL_SNAP_TOLERANCE_DEG or abs(angle - 180.0) <= CARDINAL_SNAP_TOLERANCE_DEG:
        return 0.0
    if abs(angle - 90.0) <= CARDINAL_SNAP_TOLERANCE_DEG:
        return 90.0
    return angle


def _exact_dir_for_cardinal(target_angle):
    """Return an exact XYZ unit vector for 0-deg and 90-deg targets, else None."""
    snapped = _snap_cardinal_angle(target_angle)
    if snapped == 0.0:
        return DB.XYZ(1.0, 0.0, 0.0)  # type: ignore
    if snapped == 90.0:
        return DB.XYZ(0.0, 1.0, 0.0)  # type: ignore
    return None


def _pair_perpendicular_diff(angle_a, angle_b):
    diff = abs(_normalize_180(angle_a) - _normalize_180(angle_b))
    if diff > 90.0:
        diff = 180.0 - diff
    return diff

# ---------------------------------------------------------------------------
# Geometry helpers
# ---------------------------------------------------------------------------

def _grid_midpoint(grid):
    curve = getattr(grid, "Curve", None)
    if curve is None:
        return None
    try:
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)
        return DB.XYZ((start.X + end.X) * 0.5, (start.Y + end.Y) * 0.5, (start.Z + end.Z) * 0.5)  # type: ignore
    except Exception:
        return None


def _dimension_direction_xy(dim):
    curve = getattr(dim, "Curve", None)
    if curve is None or not isinstance(curve, DB.Line):
        return None
    direction = curve.Direction
    magnitude = math.hypot(direction.X, direction.Y)
    if magnitude <= 1e-12:
        return None
    return DB.XYZ(direction.X / magnitude, direction.Y / magnitude, 0.0)  # type: ignore

# ---------------------------------------------------------------------------
# Off-axis element handling (Revit warning-based)
# ---------------------------------------------------------------------------

def _build_offaxis_records(doc):
    """Collect line elements referenced in 'slightly off axis' Revit warnings."""
    records = []
    seen_ids = set()

    OFF_AXIS_PHRASES = ("slightly off axis", "may cause inaccuracies")

    for warning in doc.GetWarnings():
        try:
            desc = warning.GetDescriptionText().lower()
        except Exception:
            continue
        if not any(p in desc for p in OFF_AXIS_PHRASES):
            continue

        try:
            failing = warning.GetFailingElements()
        except Exception:
            continue

        for elem_id in failing:
            if elem_id.IntegerValue in seen_ids:
                continue

            elem = doc.GetElement(elem_id)
            if elem is None:
                continue

            # Grids are handled separately.
            if isinstance(elem, DB.Grid):
                continue

            loc = getattr(elem, "Location", None)
            if not isinstance(loc, DB.LocationCurve):
                continue

            curve = loc.Curve
            if not isinstance(curve, DB.Line):
                continue

            angle = _line_world_angle_180(curve)
            if angle is None:
                continue

            target = _snap_cardinal_angle(_normalize_180(round(angle, ANGLE_PRECISION)))
            records.append({
                "elem": elem,
                "loc": loc,
                "curve": curve,
                "current": angle,
                "target": target,
            })
            seen_ids.add(elem_id.IntegerValue)

    return records


def _apply_element_rotations(doc, records):
    """Rotate off-axis line elements to their target cardinal angle."""
    updated = 0
    unchanged = 0
    failed = 0
    failed_items = []

    transaction = DB.Transaction(doc, TITLE + " - Elements")  # type: ignore
    transaction.Start()
    try:
        for record in records:
            elem = record["elem"]
            delta = _shortest_delta_180(record["current"], record["target"])

            if abs(delta) <= ANGLE_EPSILON:
                unchanged += 1
                continue

            repin = False
            try:
                if getattr(elem, "Pinned", False):
                    elem.Pinned = False
                    repin = True

                loc = record["loc"]
                curve = record["curve"]
                exact_dir = _exact_dir_for_cardinal(record["target"])

                if exact_dir is not None:
                    # Place the new line from start with an exact direction to
                    # eliminate any floating-point residual in the angle.
                    start = curve.GetEndPoint(0)
                    length = curve.Length
                    new_end = DB.XYZ(  # type: ignore
                        start.X + exact_dir.X * length,
                        start.Y + exact_dir.Y * length,
                        start.Z,
                    )
                    loc.Curve = DB.Line.CreateBound(start, new_end)  # type: ignore
                else:
                    start = curve.GetEndPoint(0)
                    transform = DB.Transform.CreateRotationAtPoint(DB.XYZ.BasisZ, math.radians(delta), start)  # type: ignore
                    loc.Curve = curve.CreateTransformed(transform)  # type: ignore

                updated += 1
            except Exception as ex:
                failed += 1
                eid = elem_id = getattr(getattr(elem, "Id", None), "IntegerValue", -1)
                failed_items.append((eid, str(ex)))
            finally:
                if repin:
                    try:
                        elem.Pinned = True
                    except Exception:
                        pass

        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    return updated, unchanged, failed, failed_items

# ---------------------------------------------------------------------------
# Grid angle handling
# ---------------------------------------------------------------------------

def _build_grid_records(doc):
    records = []
    non_linear = 0

    collector = DB.FilteredElementCollector(doc).OfClass(DB.Grid).WhereElementIsNotElementType()
    for grid in collector.ToElements():
        curve = getattr(grid, "Curve", None)
        line = curve if isinstance(curve, DB.Line) else None
        if line is None:
            non_linear += 1
            continue

        current_angle = _line_world_angle_180(line)
        if current_angle is None:
            non_linear += 1
            continue

        rounded_target = _normalize_180(round(current_angle, ANGLE_PRECISION))
        rounded_target = _snap_cardinal_angle(rounded_target)
        records.append({
            "grid": grid,
            "line": line,
            "current": current_angle,
            "target": rounded_target,
        })

    return records, non_linear


def _snap_near_perpendicular_targets(records):
    if len(records) < 2:
        return 0

    snapped_pairs = 0
    adjusted_indices = set()

    for i in range(len(records)):
        for j in range(i + 1, len(records)):
            angle_i = records[i]["target"]
            angle_j = records[j]["target"]
            diff = _pair_perpendicular_diff(angle_i, angle_j)
            if abs(diff - 90.0) > PERP_CANDIDATE_TOLERANCE_DEG:
                continue
            if abs(diff - 90.0) <= 1e-16:
                continue
            if i in adjusted_indices and j in adjusted_indices:
                continue

            current_i = records[i]["current"]
            current_j = records[j]["current"]

            candidate_i_target = _normalize_180(records[j]["target"] + 90.0)
            candidate_j_target = _normalize_180(records[i]["target"] + 90.0)

            move_i = abs(_shortest_delta_180(current_i, candidate_i_target))
            move_j = abs(_shortest_delta_180(current_j, candidate_j_target))

            if (j in adjusted_indices) or (move_i <= move_j and i not in adjusted_indices):
                records[i]["target"] = candidate_i_target
                adjusted_indices.add(i)
            else:
                records[j]["target"] = candidate_j_target
                adjusted_indices.add(j)

            snapped_pairs += 1

    return snapped_pairs


def _apply_grid_rotations(doc, records):
    updated = 0
    unchanged = 0
    failed = 0
    failed_items = []

    transaction = DB.Transaction(doc, TITLE + " - Grids")  # type: ignore
    transaction.Start()
    try:
        for record in records:
            grid = record["grid"]
            current_angle = record["current"]
            target_angle = _snap_cardinal_angle(_normalize_180(round(record["target"], ANGLE_PRECISION)))
            delta = _shortest_delta_180(current_angle, target_angle)

            if abs(delta) <= ANGLE_EPSILON:
                unchanged += 1
                continue

            repin = False
            try:
                if grid.Pinned:
                    grid.Pinned = False
                    repin = True

                line = record["line"]
                start = line.GetEndPoint(0)
                end = line.GetEndPoint(1)
                midpoint = DB.XYZ((start.X + end.X) * 0.5, (start.Y + end.Y) * 0.5, (start.Z + end.Z) * 0.5)  # type: ignore

                axis = DB.Line.CreateUnbound(midpoint, DB.XYZ.BasisZ)  # type: ignore
                DB.ElementTransformUtils.RotateElement(doc, grid.Id, axis, math.radians(delta))  # type: ignore

                updated += 1
            except Exception as ex:
                failed += 1
                grid_id = getattr(getattr(grid, "Id", None), "IntegerValue", -1)
                failed_items.append((grid_id, str(ex)))
            finally:
                if repin:
                    try:
                        grid.Pinned = True
                    except Exception:
                        pass

        doc.Regenerate()
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    return updated, unchanged, failed, failed_items

# ---------------------------------------------------------------------------
# Grid spacing snap (dimension-chain based, with shake)
# ---------------------------------------------------------------------------

def _collect_grid_targets(doc, view):
    if view is None:
        return []

    pending = []
    processed_grid_ids = set()

    for dim in DB.FilteredElementCollector(doc, view.Id).OfClass(DB.Dimension):  # type: ignore
        try:
            curve = dim.Curve
        except Exception:
            continue
        if not isinstance(curve, DB.Line):
            continue

        try:
            ref_array = dim.References
            if ref_array is None or ref_array.Size < 2:
                continue
        except Exception:
            continue

        refs = []
        try:
            for i in range(ref_array.Size):
                reference = ref_array.get_Item(i)
                element = doc.GetElement(reference.ElementId)
                if isinstance(element, DB.Grid):
                    refs.append(element)
        except Exception:
            continue

        if len(refs) < 2:
            continue

        axis_direction = _dimension_direction_xy(dim)
        if axis_direction is None:
            continue

        try:
            segments = dim.Segments
        except Exception:
            segments = None

        has_segments = segments is not None and segments.Size > 0

        if not has_segments and len(refs) == 2:
            try:
                single_value = dim.Value
            except Exception:
                continue
            if single_value is None:
                continue
            anchor_mid = _grid_midpoint(refs[0])
            if anchor_mid is None:
                continue
            anchor_raw = anchor_mid.X * axis_direction.X + anchor_mid.Y * axis_direction.Y
            anchor_scalar = round(anchor_raw * 304.8) / 304.8
            if refs[0].Id.IntegerValue not in processed_grid_ids and abs(anchor_scalar - anchor_raw) > 1e-12:
                processed_grid_ids.add(refs[0].Id.IntegerValue)
                pending.append((refs[0], axis_direction, anchor_scalar))
            target_scalar = anchor_scalar + round(single_value * 304.8) / 304.8
            grid = refs[1]
            if grid.Id.IntegerValue not in processed_grid_ids:
                processed_grid_ids.add(grid.Id.IntegerValue)
                pending.append((grid, axis_direction, target_scalar))
            continue

        if not has_segments:
            continue

        anchor_mid = _grid_midpoint(refs[0])
        if anchor_mid is None:
            continue
        anchor_raw = anchor_mid.X * axis_direction.X + anchor_mid.Y * axis_direction.Y
        anchor_scalar = round(anchor_raw * 304.8) / 304.8
        if refs[0].Id.IntegerValue not in processed_grid_ids and abs(anchor_scalar - anchor_raw) > 1e-12:
            processed_grid_ids.add(refs[0].Id.IntegerValue)
            pending.append((refs[0], axis_direction, anchor_scalar))
        cumulative_mm = 0.0
        assert segments is not None

        for index in range(1, len(refs)):
            try:
                segment = segments.get_Item(index - 1)
            except Exception:
                break
            cumulative_mm += round(segment.Value * 304.8)
            target_scalar = anchor_scalar + cumulative_mm / 304.8
            grid = refs[index]
            if grid.Id.IntegerValue not in processed_grid_ids:
                processed_grid_ids.add(grid.Id.IntegerValue)
                pending.append((grid, axis_direction, target_scalar))

    return pending


def _snap_grid_spacing(doc, view):
    pending = _collect_grid_targets(doc, view)
    if not pending:
        return 0, []

    moved = 0
    failures = []
    shake_ft = SHAKE_MM / 304.8

    transaction = DB.Transaction(doc, TITLE + " - Grid Spacing")  # type: ignore
    transaction.Start()
    try:
        # Pass 1: overshoot every grid to break floating-point residual.
        for grid, axis_direction, target_scalar in pending:
            grid_mid = _grid_midpoint(grid)
            if grid_mid is None:
                continue
            current_scalar = grid_mid.X * axis_direction.X + grid_mid.Y * axis_direction.Y
            delta = (target_scalar + shake_ft) - current_scalar
            move = DB.XYZ(axis_direction.X * delta, axis_direction.Y * delta, 0.0)  # type: ignore
            DB.ElementTransformUtils.MoveElement(doc, grid.Id, move)  # type: ignore

        doc.Regenerate()

        # Pass 2: move from overshoot to exact integer-mm target.
        for grid, axis_direction, target_scalar in pending:
            try:
                grid_mid = _grid_midpoint(grid)
                if grid_mid is None:
                    continue
                current_scalar = grid_mid.X * axis_direction.X + grid_mid.Y * axis_direction.Y
                delta = target_scalar - current_scalar
                if abs(delta) > 1e-12:
                    move = DB.XYZ(axis_direction.X * delta, axis_direction.Y * delta, 0.0)  # type: ignore
                    DB.ElementTransformUtils.MoveElement(doc, grid.Id, move)  # type: ignore
                moved += 1
            except Exception as ex:
                failures.append((grid.Id.IntegerValue, str(ex)))

        doc.Regenerate()
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    return moved, failures

# ---------------------------------------------------------------------------
# Spacing report
# ---------------------------------------------------------------------------

def _report_grid_spacing(doc):
    collector = DB.FilteredElementCollector(doc).OfClass(DB.Grid).WhereElementIsNotElementType()
    groups = {}

    for grid in collector.ToElements():
        curve = getattr(grid, "Curve", None)
        if not isinstance(curve, DB.Line):
            continue
        angle = _line_world_angle_180(curve)
        if angle is None:
            continue
        angle_key = int(round(_snap_cardinal_angle(angle)))

        mid = _grid_midpoint(grid)
        if mid is None:
            continue

        perp_angle = math.radians(angle_key + 90.0)
        pos = mid.X * math.cos(perp_angle) + mid.Y * math.sin(perp_angle)
        name = grid.Name if grid.Name else str(grid.Id.IntegerValue)
        groups.setdefault(angle_key, []).append((pos, name))

    if not groups:
        return

    print("")
    print("Grid Spacing")
    print("------------")
    for angle_key in sorted(groups.keys()):
        entries = sorted(groups[angle_key], key=lambda t: t[0])
        print("  {}deg grids ({} total):".format(angle_key, len(entries)))
        for i in range(1, len(entries)):
            dist_mm = (entries[i][0] - entries[i - 1][0]) * 304.8
            print("    {} -> {}: {} mm".format(entries[i - 1][1], entries[i][1], round(dist_mm)))

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    ui_doc = __revit__.ActiveUIDocument  # type: ignore
    doc = ui_doc.Document if ui_doc else None
    if doc is None:
        print("No active Revit document found.")
        return

    # --- Off-axis elements ---
    offaxis_records = _build_offaxis_records(doc)
    elem_updated, elem_unchanged, elem_failed, elem_failed_items = (0, 0, 0, [])
    if offaxis_records:
        elem_updated, elem_unchanged, elem_failed, elem_failed_items = _apply_element_rotations(doc, offaxis_records)

    # --- Grid angles ---
    grid_records, non_linear = _build_grid_records(doc)
    snapped_pairs = _snap_near_perpendicular_targets(grid_records)
    grid_updated, grid_unchanged, grid_failed, grid_failed_items = _apply_grid_rotations(doc, grid_records)

    # --- Grid spacing ---
    spacing_moved, spacing_failures = _snap_grid_spacing(doc, ui_doc.ActiveView)

    # --- Report ---
    print(TITLE)
    print("=" * len(TITLE))

    print("")
    print("Elements (off-axis warnings)")
    print("----------------------------")
    print("  Warnings found:   {}".format(len(offaxis_records)))
    print("  Fixed:            {}".format(elem_updated))
    print("  Already aligned:  {}".format(elem_unchanged))
    print("  Failed:           {}".format(elem_failed))
    if elem_failed_items:
        for eid, msg in elem_failed_items:
            print("  - Element {}: {}".format(eid, msg))

    print("")
    print("Grids (angles)")
    print("--------------")
    print("  Linear grids checked:    {}".format(len(grid_records)))
    print("  Non-linear skipped:      {}".format(non_linear))
    print("  Near-90 pairs snapped:   {}".format(snapped_pairs))
    print("  Fixed:                   {}".format(grid_updated))
    print("  Already aligned:         {}".format(grid_unchanged))
    print("  Failed:                  {}".format(grid_failed))
    if grid_failed_items:
        for gid, msg in grid_failed_items:
            print("  - Grid {}: {}".format(gid, msg))

    print("")
    print("Grids (spacing)")
    print("---------------")
    print("  Snapped to integer mm:   {}".format(spacing_moved))
    if spacing_failures:
        for gid, msg in spacing_failures:
            print("  - Grid {}: {}".format(gid, msg))

    _report_grid_spacing(doc)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print(traceback.format_exc())
