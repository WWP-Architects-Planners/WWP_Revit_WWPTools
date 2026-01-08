#! python3
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


MASS_FAMILY_PARAM = "Mass: Family"


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
                param.Set(ElementId(value))
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


def _get_mass_family_name(mass):
    try:
        symbol = mass.Symbol
        if symbol and symbol.Family:
            return symbol.Family.Name
    except Exception:
        pass

    try:
        param = mass.LookupParameter("Family")
        value = _get_param_value(param)
        if value:
            return value
    except Exception:
        pass

    try:
        return mass.Name
    except Exception:
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

    mass_name_to_index = {}
    mass_names = []
    for idx, mass in enumerate(masses):
        name = _get_mass_family_name(mass)
        mass_names.append(name)
        if name and name not in mass_name_to_index:
            mass_name_to_index[name] = idx

    updated = 0
    skipped = 0
    no_match = 0
    missing_params = 0
    type_mismatch = 0

    t = Transaction(doc, "Sync Mass Floor Parameters")
    started = False
    try:
        t.Start()
        started = True

        for mass_floor in mass_floors:
            mf_param = mass_floor.LookupParameter(MASS_FAMILY_PARAM)
            mf_name = _get_param_value(mf_param)
            if not mf_name:
                skipped += 1
                continue

            mass_index = mass_name_to_index.get(mf_name)
            if mass_index is None:
                no_match += 1
                continue

            mass = masses[mass_index]
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
        "Skipped (no Mass: Family value): {}".format(skipped),
        "Skipped (no matching Mass family): {}".format(no_match),
        "Skipped (missing/read-only params): {}".format(missing_params),
        "Skipped (storage type mismatch): {}".format(type_mismatch),
    ]
    UI.TaskDialog.Show("Sync Mass Tool", "\n".join(msg_lines))


if __name__ == "__main__":
    main()
