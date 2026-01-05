"""Sync instance parameters from Mass elements to their dependent Mass Floors.

Behavior:
- If user selection contains Mass(es): only those are processed.
- Else if selection contains Mass Floor(s): finds their owning Mass(es) and processes those.
- Else: asks to process all Masses in the model.

For each Mass -> each dependent Mass Floor:
- For every instance parameter on the Mass, if the Mass Floor has an instance parameter
  with the same name and the same storage type, and it is writable, the value is copied.

Notes:
- Works for shared parameters and project parameters alike.
- Skips read-only, missing, or incompatible parameters.
"""

from pyrevit import revit, DB, forms, script


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def _get_bic_int(bic):
    try:
        return int(bic)
    except Exception:
        try:
            return int(bic.value__)
        except Exception:
            return None


def _get_category_int(element):
    try:
        cat = element.Category
        if not cat:
            return None
        cid = cat.Id
        # Revit 2025+ has IntegerValue
        if hasattr(cid, 'IntegerValue'):
            return cid.IntegerValue
        return int(cid)
    except Exception:
        return None


def _collect_masses_in_model():
    return list(
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Mass)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def _collect_selected_elements():
    sel_ids = list(uidoc.Selection.GetElementIds())
    return [doc.GetElement(eid) for eid in sel_ids if eid]


def _is_mass(element):
    return _get_category_int(element) == _get_bic_int(DB.BuiltInCategory.OST_Mass)


def _is_mass_floor(element):
    return _get_category_int(element) == _get_bic_int(DB.BuiltInCategory.OST_MassFloor)


def _build_param_map(element):
    pmap = {}
    try:
        for p in element.Parameters:
            if not p:
                continue
            d = p.Definition
            if not d:
                continue
            name = d.Name
            if name and name not in pmap:
                pmap[name] = p
    except Exception:
        pass
    return pmap


def _copy_param_value(src_param, dst_param):
    """Copy value from src_param to dst_param.

    Returns True if destination was changed.
    """
    if not src_param or not dst_param:
        return False
    if dst_param.IsReadOnly:
        return False
    if src_param.StorageType != dst_param.StorageType:
        return False
    if not src_param.HasValue:
        return False

    st = src_param.StorageType

    try:
        if st == DB.StorageType.String:
            sval = src_param.AsString()
            if sval is None:
                sval = src_param.AsValueString()
            if sval is None:
                return False
            if dst_param.AsString() != sval:
                dst_param.Set(sval)
                return True

        elif st == DB.StorageType.Integer:
            ival = src_param.AsInteger()
            if dst_param.AsInteger() != ival:
                dst_param.Set(ival)
                return True

        elif st == DB.StorageType.Double:
            dval = src_param.AsDouble()
            if dst_param.AsDouble() != dval:
                dst_param.Set(dval)
                return True

        elif st == DB.StorageType.ElementId:
            eid = src_param.AsElementId()
            if dst_param.AsElementId() != eid:
                dst_param.Set(eid)
                return True

    except Exception:
        return False

    return False


def _get_dependent_mass_floor_ids(mass_element, mass_floor_filter):
    try:
        return list(mass_element.GetDependentElements(mass_floor_filter))
    except Exception:
        return []


# ----------------------------
# Determine scope (selection)
# ----------------------------
selected_elements = _collect_selected_elements()
selected_masses = [e for e in selected_elements if e and _is_mass(e)]
selected_mass_floors = [e for e in selected_elements if e and _is_mass_floor(e)]

mass_floor_filter = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_MassFloor)

masses_to_process = []

if selected_masses:
    masses_to_process = selected_masses

elif selected_mass_floors:
    # Find owning Mass(es) by scanning masses and matching dependents
    target_floor_ids = set([mf.Id for mf in selected_mass_floors])
    all_masses = _collect_masses_in_model()
    for m in all_masses:
        dep_ids = _get_dependent_mass_floor_ids(m, mass_floor_filter)
        for did in dep_ids:
            if did in target_floor_ids:
                masses_to_process.append(m)
                break

else:
    run_all = forms.alert(
        "No Mass or Mass Floor selected.\n\nSync ALL Masses in the model to their Mass Floors?",
        yes=True,
        no=True,
    )
    if not run_all:
        script.exit()
    masses_to_process = _collect_masses_in_model()

# De-duplicate masses
unique = {}
for m in masses_to_process:
    try:
        unique[m.Id.IntegerValue] = m
    except Exception:
        unique[m.Id] = m
masses_to_process = list(unique.values())

if not masses_to_process:
    forms.alert("No Mass elements found to process.")
    script.exit()


# ----------------------------
# Sync parameters
# ----------------------------
processed_masses = 0
found_floors = 0
updated_floors = 0
updated_params = 0

with revit.Transaction("Sync Mass parameters to Mass Floors"):
    for mass in masses_to_process:
        processed_masses += 1

        dep_floor_ids = _get_dependent_mass_floor_ids(mass, mass_floor_filter)
        if not dep_floor_ids:
            continue

        # Source parameters list (skip None storage type)
        src_params = []
        try:
            for p in mass.Parameters:
                if p and p.Definition and p.StorageType != DB.StorageType.None:
                    src_params.append(p)
        except Exception:
            src_params = []

        for floor_id in dep_floor_ids:
            mf = doc.GetElement(floor_id)
            if not mf:
                continue

            found_floors += 1
            dst_map = _build_param_map(mf)

            floor_changed = False
            for sp in src_params:
                name = None
                try:
                    name = sp.Definition.Name
                except Exception:
                    name = None

                if not name:
                    continue

                dp = dst_map.get(name)
                if not dp:
                    continue

                if _copy_param_value(sp, dp):
                    floor_changed = True
                    updated_params += 1

            if floor_changed:
                updated_floors += 1


msg = (
    "Synced Mass -> Mass Floors\n\n"
    "Masses processed: {0}\n"
    "Mass Floors found: {1}\n"
    "Mass Floors updated: {2}\n"
    "Parameters updated: {3}"
).format(processed_masses, found_floors, updated_floors, updated_params)

forms.alert(msg)
