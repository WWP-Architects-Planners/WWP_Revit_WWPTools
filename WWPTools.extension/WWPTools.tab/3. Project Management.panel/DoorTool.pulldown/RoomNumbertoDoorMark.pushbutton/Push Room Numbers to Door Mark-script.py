#! python3
from __future__ import division

import math
from pyrevit import revit, DB, script, forms


doc = revit.doc
output = script.get_output()


def get_active_phase(document):
    """Return the active view phase if available; otherwise last project phase."""
    try:
        view = document.ActiveView
        if view:
            phase_param = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
            if phase_param:
                phase_id = phase_param.AsElementId()
                if phase_id and phase_id != DB.ElementId.InvalidElementId:
                    phase = document.GetElement(phase_id)
                    if phase:
                        return phase
    except Exception:
        pass
    try:
        phases = list(document.Phases)
        if phases:
            return phases[-1]
    except Exception:
        pass
    return None


def get_room_for_door(door, phase):
    """Return the ToRoom for a door in the given phase (or active phase)."""
    try:
        if phase:
            return door.get_ToRoom(phase)
    except Exception:
        pass
    try:
        return door.ToRoom
    except Exception:
        return None


def get_room_number(room):
    if room is None:
        return None
    try:
        number = room.Number
        if number is None:
            return None
        number = str(number).strip()
        return number if number else None
    except Exception:
        return None


def get_location_point(element):
    """Return a XYZ point for element location (point or curve midpoint)."""
    try:
        loc = element.Location
    except Exception:
        return None

    if loc is None:
        return None

    if isinstance(loc, DB.LocationPoint):
        try:
            return loc.Point
        except Exception:
            return None

    if isinstance(loc, DB.LocationCurve):
        try:
            curve = loc.Curve
            if curve is None:
                return None
            return curve.Evaluate(0.5, True)
        except Exception:
            return None

    return None


def get_room_center(room):
    """Return a representative center for the room."""
    try:
        loc = room.Location
        if isinstance(loc, DB.LocationPoint):
            return loc.Point
    except Exception:
        pass

    try:
        bbox = room.get_BoundingBox(None)
        if bbox:
            return (bbox.Min + bbox.Max) * 0.5
    except Exception:
        pass

    return None


def angle_clockwise(center, point):
    """Clockwise angle around center from +X axis in the XY plane."""
    dx = point.X - center.X
    dy = point.Y - center.Y
    return -math.atan2(dy, dx)


def index_to_suffix(index):
    """Convert 0-based index to suffix letters: A, B, ... Z, AA, AB, ..."""
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    index += 1
    result = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = letters[rem] + result
    return result


def set_mark(door, value):
    param = door.LookupParameter("Mark")
    if param is None or param.IsReadOnly:
        return False
    try:
        param.Set(value)
        return True
    except Exception:
        return False


def get_door_name_tokens(door):
    """Return lowercase string of family and type names for filtering."""
    parts = []
    try:
        symbol = door.Symbol
        if symbol:
            try:
                if symbol.Family and symbol.Family.Name:
                    parts.append(symbol.Family.Name)
            except Exception:
                pass
            try:
                if symbol.Name:
                    parts.append(symbol.Name)
            except Exception:
                pass
    except Exception:
        pass
    try:
        if door.Name:
            parts.append(door.Name)
    except Exception:
        pass
    return " ".join(parts).lower().strip()


def apply_name_filter(doors, mode, raw_tokens):
    """Filter doors by family/type name tokens. mode: include|exclude."""
    if not mode or not raw_tokens:
        return doors

    tokens = [t.strip().lower() for t in raw_tokens.split(",") if t.strip()]
    if not tokens:
        return doors

    filtered = []
    for door in doors:
        hay = get_door_name_tokens(door)
        match = any(token in hay for token in tokens) if hay else False
        if mode == "include":
            if match:
                filtered.append(door)
        elif mode == "exclude":
            if not match:
                filtered.append(door)
        else:
            filtered.append(door)
    return filtered


def main():
    message = (
        "This will sync each Door Mark to its ToRoom Number.\n\n"
        "If a room has multiple doors, marks will be suffixed A, B, C...\n"
        "Doors to the same room are ordered clockwise when possible.\n\n"
        "Continue?"
    )
    if not forms.alert(message, title="Push Room Numbers to Door Mark", ok=False, yes=True, no=True):
        return

    phase = get_active_phase(doc)

    doors = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Doors)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    if not doors:
        forms.alert("No door instances found.", title="Push Room Numbers to Door Mark")
        return

    filter_choice = forms.CommandSwitchWindow.show(
        ["No filter", "Include", "Exclude"],
        message="Filter doors by family/type name?",
    )
    filter_mode = None
    filter_tokens = None
    if filter_choice in ["Include", "Exclude"]:
        filter_mode = filter_choice.lower()
        filter_tokens = forms.ask_for_string(
            prompt="Enter comma-separated keywords to {} (family/type name).".format(filter_mode),
            title="Door Name Filter",
        )

    doors = apply_name_filter(doors, filter_mode, filter_tokens)
    if not doors:
        forms.alert("No doors matched the name filter.", title="Push Room Numbers to Door Mark")
        return

    by_room_id = {}
    skipped = []

    for door in doors:
        room = get_room_for_door(door, phase)
        room_number = get_room_number(room)
        if room is None or room_number is None:
            skipped.append(door.Id)
            continue

        room_id = room.Id.IntegerValue
        entry = by_room_id.setdefault(room_id, {"room": room, "doors": [], "room_number": room_number})
        entry["doors"].append(door)

    if not by_room_id:
        forms.alert("No doors with valid ToRoom found.", title="Push Room Numbers to Door Mark")
        return

    updated = 0
    unchanged = 0
    failed = 0

    with revit.Transaction("Push Room Numbers to Door Mark"):
        for room_id, data in by_room_id.items():
            room = data["room"]
            room_number = data["room_number"]
            doors_for_room = list(data["doors"])

            center = get_room_center(room)
            if center:
                door_angles = []
                for door in doors_for_room:
                    pt = get_location_point(door)
                    if pt is None:
                        door_angles.append((None, door))
                    else:
                        door_angles.append((angle_clockwise(center, pt), door))
                sortable = [pair for pair in door_angles if pair[0] is not None]
                unsortable = [pair for pair in door_angles if pair[0] is None]
                sortable.sort(key=lambda x: x[0])
                doors_for_room = [d for _, d in sortable] + [d for _, d in unsortable]

            if len(doors_for_room) == 1:
                target_values = [room_number]
            else:
                target_values = ["{}{}".format(room_number, index_to_suffix(i)) for i in range(len(doors_for_room))]

            for door, target in zip(doors_for_room, target_values):
                param = door.LookupParameter("Mark")
                current = None
                if param is not None:
                    try:
                        current = param.AsString()
                    except Exception:
                        current = None

                if current == target:
                    unchanged += 1
                    continue

                if set_mark(door, target):
                    updated += 1
                else:
                    failed += 1

    output.print_md("## Push Room Numbers to Door Mark")
    output.print_md("- Updated: **{}**".format(updated))
    output.print_md("- Unchanged: **{}**".format(unchanged))
    output.print_md("- Failed: **{}**".format(failed))
    output.print_md("- Skipped (no ToRoom/number): **{}**".format(len(skipped)))

    forms.alert(
        "Updated: {}\nUnchanged: {}\nFailed: {}\nSkipped (no ToRoom/number): {}".format(
            updated, unchanged, failed, len(skipped)
        ),
        title="Push Room Numbers to Door Mark",
    )


if __name__ == "__main__":
    main()
