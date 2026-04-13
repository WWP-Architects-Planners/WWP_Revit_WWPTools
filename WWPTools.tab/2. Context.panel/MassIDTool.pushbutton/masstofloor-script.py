from System import Int64
import clr
import traceback

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    StorageType,
    Transaction,
)
from Autodesk.Revit import UI


def _get_doc():
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is None:
            return None
        return uidoc.Document
    except Exception:
        return None


def _collect_instances(doc, bic):
    return list(
        FilteredElementCollector(doc)
        .OfCategory(bic)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def _get_param_value(param):
    if param is None:
        return None
    stype = param.StorageType
    if stype == StorageType.String:
        return param.AsString()
    if stype == StorageType.Integer:
        return param.AsInteger()
    if stype == StorageType.Double:
        return param.AsDouble()
    if stype == StorageType.ElementId:
        return param.AsElementId()
    return None


def _set_param_value(param, value):
    if param is None or param.IsReadOnly:
        return False

    stype = param.StorageType
    try:
        if stype == StorageType.String:
            param.Set("" if value is None else str(value))
            return True
        if stype == StorageType.Integer:
            if value is None:
                return False
            param.Set(int(value))
            return True
        if stype == StorageType.Double:
            if value is None:
                return False
            param.Set(float(value))
            return True
        if stype == StorageType.ElementId:
            if isinstance(value, ElementId):
                param.Set(value)
                return True
            if isinstance(value, int):
                param.Set(ElementId(Int64(value)))
                return True
    except Exception:
        return False

    return False


def _build_param_map(element):
    param_map = {}
    for param in element.Parameters:
        try:
            name = param.Definition.Name
        except Exception:
            continue
        if name and name not in param_map:
            param_map[name] = param
    return param_map


def _iter_instance_params(element):
    for param in element.Parameters:
        try:
            _ = param.Definition.Name
        except Exception:
            continue
        yield param


def _get_parent_mass_id(mass_floor):
    """Return the ElementId of the mass that owns this mass floor.

    OST_MassFloor elements are MassLevelData instances in Revit 2024-2027.
    OwningMassId is the primary API (stable since Revit 2014).
    Falls back to the HOST_ID_PARAM built-in for any version where the
    property name differs.
    """
    # Primary: MassLevelData.OwningMassId (Revit 2014+)
    try:
        eid = mass_floor.OwningMassId
        if eid is not None and eid != ElementId.InvalidElementId:
            return eid
    except Exception:
        pass

    # Fallback: HOST_ID_PARAM built-in parameter
    try:
        from Autodesk.Revit.DB import BuiltInParameter
        p = mass_floor.get_Parameter(BuiltInParameter.HOST_ID_PARAM)
        if p is not None:
            eid = p.AsElementId()
            if eid is not None and eid != ElementId.InvalidElementId:
                return eid
    except Exception:
        pass

    return None


def main():
    doc = _get_doc()
    if doc is None:
        UI.TaskDialog.Show("Sync Mass Tool", "No active Revit document found.")
        return

    masses = _collect_instances(doc, BuiltInCategory.OST_Mass)
    mass_floors = _collect_instances(doc, BuiltInCategory.OST_MassFloor)

    if not masses:
        UI.TaskDialog.Show("Sync Mass Tool", "No Mass elements found.")
        return
    if not mass_floors:
        UI.TaskDialog.Show("Sync Mass Tool", "No Mass Floors found.")
        return

    mass_id_to_mass = {mass.Id: mass for mass in masses}

    updated = 0
    skipped = 0
    no_match = 0
    missing_params = 0
    type_mismatch = 0
    synced_param_counts = {}  # param name -> number of floors it was written to

    t = Transaction(doc, "Sync Mass Floor Parameters")
    started = False
    try:
        t.Start()
        started = True

        for mass_floor in mass_floors:
            parent_mass_id = _get_parent_mass_id(mass_floor)
            if parent_mass_id is None:
                skipped += 1
                continue

            mass = mass_id_to_mass.get(parent_mass_id)
            if mass is None:
                no_match += 1
                continue

            mass_floor_params = _build_param_map(mass_floor)

            wrote_any = False
            for src_param in _iter_instance_params(mass):
                name = src_param.Definition.Name
                target_param = mass_floor_params.get(name)
                if target_param is None:
                    continue

                if target_param.StorageType != src_param.StorageType:
                    type_mismatch += 1
                    continue

                value = _get_param_value(src_param)
                if _set_param_value(target_param, value):
                    wrote_any = True
                    synced_param_counts[name] = synced_param_counts.get(name, 0) + 1

            if wrote_any:
                updated += 1
            else:
                missing_params += 1

        t.Commit()
    except Exception as exc:
        if started:
            try:
                t.RollBack()
            except Exception:
                pass
        UI.TaskDialog.Show(
            "Sync Mass Tool - Error",
            "{}\n\n{}".format(exc, traceback.format_exc()),
        )
        return

    msg_lines = [
        "Updated: {}".format(updated),
        "Skipped (no parent mass link): {}".format(skipped),
        "Skipped (parent mass not in model): {}".format(no_match),
        "Skipped (missing/read-only params): {}".format(missing_params),
        "Skipped (storage type mismatch): {}".format(type_mismatch),
        "",
        "Parameters synced (name: floors written):",
    ]
    if synced_param_counts:
        for name in sorted(synced_param_counts):
            msg_lines.append("  {}: {}".format(name, synced_param_counts[name]))
    else:
        msg_lines.append("  (none)")
    UI.TaskDialog.Show("Sync Mass Tool", "\n".join(msg_lines))


if __name__ == "__main__":
    main()
