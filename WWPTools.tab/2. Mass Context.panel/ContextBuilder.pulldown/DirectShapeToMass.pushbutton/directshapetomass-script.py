#! python3
import clr
import datetime
import os
import re
import sys
import traceback

from System.Collections.Generic import List

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit import DB, UI


TITLE = "DirectShape To Mass"
SOURCE_APP_IDS = ("WWPTools.WebContextBuilder", "WWPTools.BuildingImporter")
MAX_FAILURES_IN_REPORT = 30
DEBUG_LOG_FILE_NAME = "DirectShapeToMass.debug.log"
_DEBUG_LOG_BUFFER = []


script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.append(lib_path)

import WWP_uiUtils as ui


def _get_uidoc():
    try:
        return __revit__.ActiveUIDocument
    except Exception:
        return None


def _get_doc():
    uidoc = _get_uidoc()
    return uidoc.Document if uidoc is not None else None


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


def _get_comment(element):
    try:
        parameter = element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if parameter:
            value = parameter.AsString()
            if value:
                return value
    except Exception:
        pass

    try:
        parameter = element.LookupParameter("Comments")
        if parameter:
            value = parameter.AsString()
            if value:
                return value
    except Exception:
        pass

    return ""


class _OverwriteFamilyLoadOptions(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        try:
            overwriteParameterValues.Value = True
        except Exception:
            pass
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        try:
            source.Value = DB.FamilySource.Family
        except Exception:
            pass
        try:
            overwriteParameterValues.Value = True
        except Exception:
            pass
        return True


def _make_safe_name(text, fallback):
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", text or "")
    safe = safe.strip("_.")
    return safe or fallback


def _find_mass_family_template(app):
    template_root = getattr(app, "FamilyTemplatePath", None)
    if not template_root or not os.path.isdir(template_root):
        return None

    candidates = []
    for root, _, files in os.walk(template_root):
        for file_name in files:
            lowered_name = file_name.lower()
            if not lowered_name.endswith(".rft"):
                continue
            if "mass" not in lowered_name:
                continue
            full_path = os.path.join(root, file_name)
            lowered_path = full_path.lower()
            score = 100
            if "conceptual mass" in lowered_path:
                score = 0
            elif "metric" in lowered_name and "mass" in lowered_name:
                score = 10
            elif lowered_name == "mass.rft":
                score = 20
            candidates.append((score, len(full_path), full_path))

    if not candidates:
        return None

    candidates.sort()
    return candidates[0][2]


def _ensure_cache_dir():
    cache_dir = os.path.join(
        os.getenv("APPDATA") or os.path.expanduser("~"),
        "pyRevit",
        "WWPTools",
        "DirectShapeToMassCache",
    )
    if not os.path.isdir(cache_dir):
        os.makedirs(cache_dir)
    return cache_dir


def _debug_log_path():
    return os.path.join(_ensure_cache_dir(), DEBUG_LOG_FILE_NAME)


def _reset_debug_log():
    global _DEBUG_LOG_BUFFER
    _DEBUG_LOG_BUFFER = ["[{}] {} debug session started".format(datetime.datetime.now().isoformat(), TITLE)]
    return _debug_log_path()


def _log_debug(message):
    global _DEBUG_LOG_BUFFER
    _DEBUG_LOG_BUFFER.append("[{}] {}".format(datetime.datetime.now().isoformat(), message))


def _read_debug_log():
    try:
        return "\n".join(_DEBUG_LOG_BUFFER).strip()
    except Exception:
        return ""


def _flush_debug_log():
    log_path = _debug_log_path()
    try:
        with open(log_path, "w") as log_file:
            log_file.write(_read_debug_log() + "\n")
    except Exception:
        pass
    return log_path


def _is_target_category(element):
    try:
        category_id_value = element.Category.Id.IntegerValue
        return category_id_value in (int(DB.BuiltInCategory.OST_Mass), int(DB.BuiltInCategory.OST_GenericModel))
    except Exception:
        return False


def _is_supported_import(element, allow_any_directshape=False):
    if element is None or not isinstance(element, DB.DirectShape):
        return False
    if not _is_target_category(element):
        return False
    if allow_any_directshape:
        return True

    app_id = (getattr(element, "ApplicationId", None) or "").strip()
    if app_id in SOURCE_APP_IDS:
        return True

    comment = _get_comment(element)
    return comment.startswith("Web Context Builder |") or comment.startswith("OSM feature:")


def _iter_solids(geometry_element):
    solids = []
    if geometry_element is None:
        return solids

    for geometry_object in geometry_element:
        if isinstance(geometry_object, DB.Solid):
            try:
                if geometry_object.Volume > 1e-9 and geometry_object.Faces.Size > 0:
                    solids.append(geometry_object)
            except Exception:
                pass
        elif isinstance(geometry_object, DB.GeometryInstance):
            try:
                solids.extend(_iter_solids(geometry_object.GetInstanceGeometry()))
            except Exception:
                pass

    return solids


def _get_element_solids(element):
    options = DB.Options()
    try:
        options.IncludeNonVisibleObjects = True
    except Exception:
        pass

    geometry = element.get_Geometry(options)
    solids = _iter_solids(geometry)
    solids.sort(key=lambda solid: solid.Volume, reverse=True)
    return solids


def _get_candidate_elements(doc, uidoc):
    selected = []
    try:
        for element_id in uidoc.Selection.GetElementIds():
            element = doc.GetElement(element_id)
            if _is_supported_import(element, allow_any_directshape=True):
                selected.append(element)
    except Exception:
        selected = []

    if selected:
        _log_debug("Using {} selected DirectShape element(s).".format(len(selected)))
        return selected, "Current selection"

    imported = []
    for element in DB.FilteredElementCollector(doc).OfClass(DB.DirectShape):
        if _is_supported_import(element, allow_any_directshape=False):
            imported.append(element)

    _log_debug("Using {} imported DirectShape element(s).".format(len(imported)))
    return imported, "Imported Web Context / Building Importer direct shapes"


def _load_family_into_project(doc, family_path, family_name):
    _log_debug("Loading family into project. path='{}' expected_name='{}'".format(family_path, family_name))
    _log_debug("Project doc.IsModifiable before load: {}".format(getattr(doc, "IsModifiable", "<unknown>")))
    before_family_ids = set()
    try:
        for candidate in DB.FilteredElementCollector(doc).OfClass(DB.Family):
            before_family_ids.add(candidate.Id.IntegerValue)
    except Exception:
        pass
    _log_debug("Project family count before load: {}".format(len(before_family_ids)))

    load_error = None
    family = None
    transaction = DB.Transaction(doc, "Load Converted Mass Family")
    started = False
    try:
        transaction.Start()
        started = True
        _log_debug("Project doc.IsModifiable inside load transaction: {}".format(getattr(doc, "IsModifiable", "<unknown>")))

        load_attempts = [
            ("doc.LoadFamily(path)", lambda: doc.LoadFamily(family_path)),
            ("doc.LoadFamily(path, options)", lambda: doc.LoadFamily(family_path, _OverwriteFamilyLoadOptions())),
        ]
        for label, attempt in load_attempts:
            try:
                loaded = attempt()
                _log_debug("{} returned type='{}' value='{}'".format(label, type(loaded).__name__, loaded))
                if isinstance(loaded, DB.Family):
                    family = loaded
                    break
                if isinstance(loaded, (list, tuple)):
                    for item in loaded:
                        if isinstance(item, DB.Family):
                            family = item
                            break
                    if family is not None:
                        break
            except Exception as ex:
                load_error = str(ex)
                _log_debug("{} raised: {}".format(label, load_error))

        transaction.Commit()
    except Exception:
        if started:
            try:
                transaction.RollBack()
            except Exception:
                pass
        raise

    try:
        if family is None:
            family_name_stem = os.path.splitext(os.path.basename(family_path))[0]
            exact_matches = []
            stem_matches = []
            prefix_matches = []
            new_families = []

            for candidate in DB.FilteredElementCollector(doc).OfClass(DB.Family):
                try:
                    candidate_name = candidate.Name or ""
                    if candidate.Id.IntegerValue not in before_family_ids:
                        new_families.append(candidate)
                    if candidate_name == family_name:
                        exact_matches.append(candidate)
                    elif candidate_name == family_name_stem:
                        stem_matches.append(candidate)
                    elif candidate_name.startswith(family_name) or candidate_name.startswith(family_name_stem):
                        prefix_matches.append(candidate)
                except Exception:
                    pass

            _log_debug(
                "Family lookup after load: exact_matches={} stem_matches={} prefix_matches={} new_families={} stem='{}'".format(
                    len(exact_matches),
                    len(stem_matches),
                    len(prefix_matches),
                    len(new_families),
                    family_name_stem,
                )
            )
            if new_families:
                _log_debug("New family names: {}".format(", ".join(sorted([(candidate.Name or "<unnamed>") for candidate in new_families][:20]))))

            if exact_matches:
                family = exact_matches[0]
            elif stem_matches:
                family = stem_matches[0]
            elif len(new_families) == 1:
                family = new_families[0]
            elif prefix_matches:
                family = prefix_matches[0]
    except Exception:
        pass

    if family is not None:
        try:
            _log_debug("Resolved loaded family: id={} name='{}'".format(family.Id.IntegerValue, family.Name or "<unnamed>"))
        except Exception:
            _log_debug("Resolved loaded family object.")
    else:
        _log_debug("Failed to resolve loaded family object after load.")

    return family, load_error


def _place_family_instance(doc, family, comment_text):
    try:
        _log_debug("Placing family instance for family id={} name='{}'".format(family.Id.IntegerValue, family.Name or "<unnamed>"))
    except Exception:
        _log_debug("Placing family instance for unresolved family metadata.")
    symbol_ids = list(family.GetFamilySymbolIds())
    if not symbol_ids:
        raise Exception("The loaded family contains no types.")

    symbol = doc.GetElement(symbol_ids[0])
    if symbol is None:
        raise Exception("Failed to access the loaded family type.")

    transaction = DB.Transaction(doc, "Place Converted Mass Family")
    started = False
    try:
        transaction.Start()
        started = True

        try:
            if hasattr(symbol, "IsActive") and not symbol.IsActive:
                symbol.Activate()
                doc.Regenerate()
        except Exception:
            pass

        creation_attempts = [
            lambda: doc.Create.NewFamilyInstance(DB.XYZ(0.0, 0.0, 0.0), symbol, DB.Structure.StructuralType.NonStructural),
            lambda: doc.Create.NewFamilyInstance(DB.XYZ(0.0, 0.0, 0.0), symbol, doc.ActiveView),
            lambda: doc.Create.NewFamilyInstance(DB.XYZ(0.0, 0.0, 0.0), symbol, doc.GetElement(doc.ActiveView.GenLevel.Id), DB.Structure.StructuralType.NonStructural),
        ]
        errors = []
        for attempt in creation_attempts:
            try:
                instance = attempt()
                if instance is not None:
                    _set_comment(instance, comment_text)
                    transaction.Commit()
                    _log_debug("Placed family instance id={}".format(instance.Id.IntegerValue))
                    return instance
            except Exception as ex:
                errors.append(str(ex))
                _log_debug("Family placement attempt failed: {}".format(str(ex)))

        raise Exception("Failed to place the generated mass family instance. {}".format("; ".join(errors[:5])))
    except Exception:
        if started:
            try:
                transaction.RollBack()
            except Exception:
                pass
        raise


def _add_material_instance_param(family_doc, freeform_elements):
    """Add a 'Material' instance parameter to the family and associate it with the primary freeform geometry."""
    if not freeform_elements:
        return

    family_manager = family_doc.FamilyManager
    mat_family_param = None

    # Revit 2022+ API uses ForgeTypeId-based overload
    try:
        mat_family_param = family_manager.AddParameter(
            "Material",
            DB.GroupTypeId.Materials,
            DB.SpecTypeId.Reference.Material,
            True,  # instance parameter
        )
    except Exception:
        pass

    # Fallback for pre-2022 Revit (ParameterType enum)
    if mat_family_param is None:
        try:
            mat_family_param = family_manager.AddParameter(
                "Material",
                DB.BuiltInParameterGroup.PG_MATERIALS,
                DB.ParameterType.Material,
                True,  # instance parameter
            )
        except Exception:
            pass

    if mat_family_param is None:
        _log_debug("Could not add Material family instance parameter - skipping association.")
        return

    # Associate the largest solid's freeform element material parameter to the family parameter
    freeform = freeform_elements[0]
    mat_elem_param = freeform.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM)
    if mat_elem_param is None:
        mat_elem_param = freeform.LookupParameter("Material")

    if mat_elem_param is not None:
        try:
            family_manager.AssociateElementParameterToFamilyParameter(mat_elem_param, mat_family_param)
            _log_debug("Associated freeform material parameter to family instance parameter 'Material'.")
        except Exception as ex:
            _log_debug("Could not associate material parameter: {}".format(ex))
    else:
        _log_debug("No material parameter found on freeform element - association skipped.")


def _convert_element_to_mass_family(doc, element, solids):
    _log_debug("Converting DirectShape element id={} app_id='{}' app_data='{}' solids={}".format(
        element.Id.IntegerValue,
        (getattr(element, "ApplicationId", None) or ""),
        (getattr(element, "ApplicationDataId", None) or ""),
        len(solids),
    ))
    app = getattr(doc, "Application", None)
    if app is None:
        raise Exception("Unable to access the Revit application.")

    template_path = _find_mass_family_template(app)
    if not template_path:
        template_root = getattr(app, "FamilyTemplatePath", None) or "<not set>"
        raise Exception(
            "No mass family template was found under Revit's Family Template Path.\n"
            "Current path: {}\n"
            "Set Revit's Family Template Path to a location that contains a Conceptual Mass template and try again."
            .format(template_root)
        )
    _log_debug("Using mass family template: {}".format(template_path))

    source_key = (getattr(element, "ApplicationDataId", None) or "").strip()
    if not source_key:
        source_key = "Element_{}".format(element.Id.IntegerValue)
    family_name = _make_safe_name("WWP_ConvertedMass_{}".format(source_key), "WWP_ConvertedMass_{}".format(element.Id.IntegerValue))

    family_doc = app.NewFamilyDocument(template_path)
    if family_doc is None:
        raise Exception("Failed to open a new conceptual mass family document.")
    _log_debug("Opened new family document for '{}'.".format(family_name))

    family_path = os.path.join(_ensure_cache_dir(), "{}.rfa".format(family_name))
    transaction = None
    try:
        transaction = DB.Transaction(family_doc, "Create Converted Mass")
        transaction.Start()

        owner_family = getattr(family_doc, "OwnerFamily", None)
        if owner_family is not None:
            try:
                owner_family.Name = family_name
            except Exception:
                pass

        freeform_cls = getattr(DB, "FreeFormElement", None)
        if freeform_cls is None:
            raise Exception("This Revit version does not support FreeFormElement creation.")

        source_comment = _get_comment(element)
        created_freeforms = []
        for solid in solids:
            freeform = freeform_cls.Create(family_doc, solid)
            if freeform is None:
                raise Exception("Revit did not create a freeform element in the mass family.")
            _set_comment(freeform, source_comment)
            created_freeforms.append(freeform)

        _add_material_instance_param(family_doc, created_freeforms)

        transaction.Commit()
        transaction = None

        save_options = DB.SaveAsOptions()
        save_options.OverwriteExistingFile = True
        family_doc.SaveAs(family_path, save_options)
        _log_debug("Saved generated family to '{}'.".format(family_path))
    except Exception:
        if transaction is not None:
            try:
                transaction.RollBack()
            except Exception:
                pass
        _log_debug("Conversion failed before family load:\n{}".format(traceback.format_exc()))
        raise
    finally:
        try:
            family_doc.Close(False)
        except Exception:
            pass
        _log_debug("Closed family document for '{}'.".format(family_name))

    family, load_error = _load_family_into_project(doc, family_path, family_name)
    if family is None:
        if load_error:
            raise Exception("Failed to load the generated mass family into the project. {}".format(load_error))
        raise Exception("Failed to load the generated mass family into the project.")

    comment_text = "Converted from DirectShape {}{}".format(
        element.Id.IntegerValue,
        " | {}".format(_get_comment(element)) if _get_comment(element) else "",
    )
    return _place_family_instance(doc, family, comment_text)


def _preview_text(scope_label, elements):
    lines = [
        "Scope: {}".format(scope_label),
        "DirectShape buildings found: {}".format(len(elements)),
        "Output: one conceptual mass family per source DirectShape",
        "Original DirectShapes will be kept.",
        "",
        "Preview of first 20 source elements:",
    ]

    for element in elements[:20]:
        lines.append(
            "Element {} | {} | {}".format(
                element.Id.IntegerValue,
                (getattr(element, "ApplicationDataId", None) or "<no app data id>"),
                _get_comment(element) or "<no comment>",
            )
        )

    if len(elements) > 20:
        lines.append("...")

    return "\n".join(lines)


def _summarize_results(created_count, failures):
    lines = [
        "Mass families created: {}".format(created_count),
        "Failures: {}".format(len(failures)),
    ]

    if failures:
        lines.append("")
        lines.append("Failures (first {}):".format(MAX_FAILURES_IN_REPORT))
        for failure in failures[:MAX_FAILURES_IN_REPORT]:
            lines.append("{} | {}".format(failure["id"], failure["reason"]))

    return "\n".join(lines)


def main():
    uidoc = _get_uidoc()
    doc = _get_doc()
    if uidoc is None or doc is None:
        UI.TaskDialog.Show(TITLE, "No active Revit document found.")
        return

    log_path = _reset_debug_log()
    _log_debug("Active document title='{}' version='{}'".format(getattr(doc, "Title", "<unknown>"), getattr(doc.Application, "VersionNumber", "<unknown>")))

    elements, scope_label = _get_candidate_elements(doc, uidoc)
    if not elements:
        UI.TaskDialog.Show(TITLE, "No matching DirectShape buildings were found.\nSelect DirectShape buildings, or run this after Web Context Builder / Building Importer.")
        return

    if not ui.uiUtils_show_text_report(
        "{} - Preview".format(TITLE),
        _preview_text(scope_label, elements),
        ok_text="Convert",
        cancel_text="Cancel",
        width=760,
        height=560,
    ):
        return

    created_instances = []
    failures = []

    for element in elements:
        try:
            solids = _get_element_solids(element)
            if not solids:
                raise Exception("No valid solids were found in the DirectShape geometry.")
            _log_debug("Element {} yielded {} solid(s).".format(element.Id.IntegerValue, len(solids)))

            instance = _convert_element_to_mass_family(doc, element, solids)
            if instance is not None:
                created_instances.append(instance)
        except Exception as ex:
            _log_debug("Element {} failed: {}\n{}".format(element.Id.IntegerValue, str(ex), traceback.format_exc()))
            failures.append({"id": "element/{}".format(element.Id.IntegerValue), "reason": str(ex)})

    try:
        if created_instances:
            selected_ids = List[DB.ElementId]()
            for instance in created_instances:
                selected_ids.Add(instance.Id)
            uidoc.Selection.SetElementIds(selected_ids)
    except Exception:
        pass

    result_text = _summarize_results(len(created_instances), failures)
    if failures:
        log_path = _flush_debug_log()
        result_text += "\n\nDetailed debug log:\n{}".format(log_path)

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
