# pyright: reportGeneralTypeIssues=false

import math
import traceback

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit import DB
TITLE = "Fix Grid Angles"
ANGLE_PRECISION = 2
ANGLE_EPSILON = 1e-16
PERP_CANDIDATE_TOLERANCE_DEG = 0.2
CARDINAL_SNAP_TOLERANCE_DEG = 0.01
SHAKE_MM = 500.0  # Temporary overshoot in mm to force Revit to commit position cleanly.


def _normalize_180(angle_degrees):
    value = angle_degrees % 180.0
    if value < 0.0:
        value += 180.0
    if abs(value - 180.0) <= ANGLE_EPSILON:
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


def _unit_vector_xy_from_angle_180(angle_degrees):
    snapped = _snap_cardinal_angle(angle_degrees)
    if snapped == 0.0:
        return 1.0, 0.0
    if snapped == 90.0:
        return 0.0, 1.0

    radians_value = math.radians(snapped)
    x = math.cos(radians_value)
    y = math.sin(radians_value)

    # Remove tiny floating noise close to axis directions.
    if abs(x) <= 1e-14:
        x = 0.0
    if abs(y) <= 1e-14:
        y = 0.0
    return x, y


def _pair_perpendicular_diff(angle_a, angle_b):
    diff = abs(_normalize_180(angle_a) - _normalize_180(angle_b))
    if diff > 90.0:
        diff = 180.0 - diff
    return diff


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
        records.append(
            {
                "grid": grid,
                "line": line,
                "current": current_angle,
                "target": rounded_target,
            }
        )

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

            if abs(diff - 90.0) <= ANGLE_EPSILON:
                continue

            # Keep changes stable: only snap each grid once in this pass.
            if i in adjusted_indices and j in adjusted_indices:
                continue

            current_i = records[i]["current"]
            current_j = records[j]["current"]

            # In 0..180 orientation space, perpendicular direction is +90 mod 180.
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

    transaction = DB.Transaction(doc, TITLE)  # type: ignore
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
                grid_id = getattr(getattr(grid, "Id", None), "IntegerValue", None)
                if grid_id is None and getattr(grid, "Id", None) is not None:
                    try:
                        grid_id = int(grid.Id.Value)
                    except Exception:
                        grid_id = -1
                failed_items.append((grid_id if grid_id is not None else -1, str(ex)))
            finally:
                if repin:
                    try:
                        grid.Pinned = True
                    except Exception:
                        pass

        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    return updated, unchanged, failed, failed_items


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


def _curve_direction_xy(grid):
    curve = getattr(grid, "Curve", None)
    if curve is None or not isinstance(curve, DB.Line):
        return None
    direction = curve.Direction
    magnitude = math.hypot(direction.X, direction.Y)
    if magnitude <= 1e-12:
        return None
    return DB.XYZ(direction.X / magnitude, direction.Y / magnitude, 0.0)  # type: ignore


def _move_grid_along_axis(doc, grid, axis_direction, delta_distance):
    if abs(delta_distance) <= 1e-12:
        return False
    move = DB.XYZ(axis_direction.X * delta_distance, axis_direction.Y * delta_distance, 0.0)  # type: ignore
    DB.ElementTransformUtils.MoveElement(doc, grid.Id, move)  # type: ignore
    return True


def _collect_grid_targets(doc, view):
    """Return a list of (grid, axis_direction, target_scalar_ft) for every grid
    referenced by a dimension in *view* whose position should be snapped to the
    nearest integer millimetre."""
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

        # Simple two-reference dimension: use dim.Value directly.
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


def _snap_grid_chain_dimensions_to_integers(doc, view):
    pending = _collect_grid_targets(doc, view)
    if not pending:
        return 0, []

    moved = 0
    failures = []
    shake_ft = SHAKE_MM / 304.8

    transaction = DB.Transaction(doc, "Snap Grid Chain Distances")  # type: ignore
    transaction.Start()
    try:
        # Pass 1: overshoot every grid past its target so Revit releases the
        # original floating-point position entirely.
        for grid, axis_direction, target_scalar in pending:
            grid_mid = _grid_midpoint(grid)
            if grid_mid is None:
                continue
            current_scalar = grid_mid.X * axis_direction.X + grid_mid.Y * axis_direction.Y
            delta = (target_scalar + shake_ft) - current_scalar
            move = DB.XYZ(axis_direction.X * delta, axis_direction.Y * delta, 0.0)  # type: ignore
            DB.ElementTransformUtils.MoveElement(doc, grid.Id, move)  # type: ignore

        # Commit the overshoot so Revit re-computes positions from scratch.
        doc.Regenerate()

        # Pass 2: move each grid from its overshooted position back to the
        # exact integer-mm target.  Delta is computed from the live position
        # after regeneration, not from a fixed -500 mm offset.
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


def _report_grid_spacing(doc):
    """Group grids by orientation and print spacing between consecutive grids."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.Grid).WhereElementIsNotElementType()
    groups = {}  # rounded_angle -> list of (position, name)

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

        # Position along the axis perpendicular to the grid direction.
        perp_angle = math.radians(angle_key + 90.0)
        perp_x = math.cos(perp_angle)
        perp_y = math.sin(perp_angle)
        pos = mid.X * perp_x + mid.Y * perp_y

        name = grid.Name if grid.Name else str(grid.Id.IntegerValue)
        groups.setdefault(angle_key, []).append((pos, name))

    if not groups:
        return

    print("")
    print("Grid Spacing")
    print("------------")
    for angle_key in sorted(groups.keys()):
        entries = sorted(groups[angle_key], key=lambda t: t[0])
        label = "{}deg grids".format(angle_key)
        print("  {} ({} total):".format(label, len(entries)))
        for i in range(1, len(entries)):
            dist_mm = (entries[i][0] - entries[i - 1][0]) * 304.8
            print("    {} -> {}: {} mm".format(entries[i - 1][1], entries[i][1], round(dist_mm)))


def main():
    ui_doc = __revit__.ActiveUIDocument  # type: ignore
    doc = ui_doc.Document if ui_doc else None
    if doc is None:
        print("No active Revit document found.")
        return

    records, non_linear = _build_grid_records(doc)
    total_linear = len(records)
    if total_linear == 0:
        print("No linear grid lines found. Non-linear grids skipped: {}".format(non_linear))
        return

    snapped_pairs = _snap_near_perpendicular_targets(records)
    updated, unchanged, failed, failed_items = _apply_grid_rotations(doc, records)
    snapped_grid_chain_count, snap_failures = _snap_grid_chain_dimensions_to_integers(doc, ui_doc.ActiveView)

    print(TITLE)
    print("=" * len(TITLE))
    print("Linear grids checked: {}".format(total_linear))
    print("Non-linear grids skipped: {}".format(non_linear))
    print("Near-90 pairs snapped: {}".format(snapped_pairs))
    print("Updated grids: {}".format(updated))
    print("Already aligned: {}".format(unchanged))
    print("Failed updates: {}".format(failed))
    print("Grid chain refs snapped: {}".format(snapped_grid_chain_count))
    if failed_items:
        print("Failure details:")
        for grid_id, message in failed_items:
            print("- Grid {}: {}".format(grid_id, message))
    if snap_failures:
        print("Chain snap failures:")
        for grid_id, message in snap_failures:
            print("- Grid {}: {}".format(grid_id, message))
    _report_grid_spacing(doc)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print(traceback.format_exc())
