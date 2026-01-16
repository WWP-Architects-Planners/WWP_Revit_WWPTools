#! python3
import traceback

from Autodesk.Revit import DB

import WWP_uiUtils as ui
from UIutility import UIform


TITLE = "Duplicate Views"
DESCRIPTION = "Replace String"
PARAM_VIEW_SUBCATEGORY = "View Subcategory"


def _clean_name(name):
    if not name:
        return ""
    return name.replace("{", "").replace("}", "")


def _collect_selected_views(uidoc, doc):
    views = []
    skipped = []
    for element_id in uidoc.Selection.GetElementIds():
        element = doc.GetElement(element_id)
        if isinstance(element, DB.View) and not element.IsTemplate:
            views.append(element)
        else:
            skipped.append(element)
    return views, skipped


def _build_existing_names(doc):
    names = set()
    for view in DB.FilteredElementCollector(doc).OfClass(DB.View):
        try:
            names.add(view.Name)
        except Exception:
            pass
    return names


def _ensure_unique_name(existing, base_name):
    if base_name not in existing:
        existing.add(base_name)
        return base_name
    index = 2
    while True:
        candidate = "{} ({})".format(base_name, index)
        if candidate not in existing:
            existing.add(candidate)
            return candidate
        index += 1


def _param_to_string(param):
    if param is None:
        return ""
    try:
        if param.StorageType == DB.StorageType.String:
            return param.AsString() or ""
        value = param.AsValueString()
        if value is not None:
            return value
        if param.StorageType == DB.StorageType.Integer:
            return str(param.AsInteger())
        if param.StorageType == DB.StorageType.Double:
            return str(param.AsDouble())
        if param.StorageType == DB.StorageType.ElementId:
            return str(param.AsElementId().IntegerValue)
    except Exception:
        return ""
    return ""


def _set_view_subcategory(view, suffix, errors):
    if not suffix:
        return
    param = view.LookupParameter(PARAM_VIEW_SUBCATEGORY)
    if param is None:
        errors.append("Missing parameter '{}' on view '{}'.".format(PARAM_VIEW_SUBCATEGORY, view.Name))
        return
    if param.IsReadOnly:
        errors.append("Parameter '{}' is read-only on view '{}'.".format(PARAM_VIEW_SUBCATEGORY, view.Name))
        return
    if param.StorageType != DB.StorageType.String:
        errors.append("Parameter '{}' is not a text parameter on view '{}'.".format(PARAM_VIEW_SUBCATEGORY, view.Name))
        return
    current = _param_to_string(param)
    try:
        param.Set("{}{}".format(current, suffix))
    except Exception as exc:
        errors.append("Failed to set '{}' on view '{}': {}".format(PARAM_VIEW_SUBCATEGORY, view.Name, exc))


def _duplicate_option_from_value(value):
    mapping = {
        "Duplicate": DB.ViewDuplicateOption.Duplicate,
        "AsDependent": DB.ViewDuplicateOption.AsDependent,
        "WithDetailing": DB.ViewDuplicateOption.WithDetailing,
    }
    return mapping.get(value, DB.ViewDuplicateOption.WithDetailing)


def _build_form_result():
    options = [
        ("Duplicate View", "Duplicate"),
        ("Duplicate as Dependent", "AsDependent"),
        ("Duplicate with Details", "WithDetailing"),
    ]
    return UIform(
        title=TITLE,
        description=DESCRIPTION,
        prefix_default="",
        suffix_default="-copy",
        duplicate_options=options,
        default_index=2,
        button_text="Set Values",
    )


def main():
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None:
        return
    doc = uidoc.Document

    views, skipped = _collect_selected_views(uidoc, doc)
    if not views:
        ui.uiUtils_alert("Select at least one view.", title=TITLE)
        return

    form_result = _build_form_result()
    if not form_result:
        return

    prefix = form_result.get("prefix", "")
    suffix = form_result.get("suffix", "")
    option_value = form_result.get("duplicate_option")
    duplicate_option = _duplicate_option_from_value(option_value)

    existing_names = _build_existing_names(doc)
    errors = []
    duplicated = []

    transaction = DB.Transaction(doc, "Duplicate Views")
    transaction.Start()
    try:
        for view in views:
            try:
                new_id = view.Duplicate(duplicate_option)
                new_view = doc.GetElement(new_id)
            except Exception as exc:
                errors.append("Failed to duplicate '{}': {}".format(view.Name, exc))
                continue

            clean_name = _clean_name(view.Name)
            target_name = _ensure_unique_name(existing_names, "{}{}{}".format(prefix, clean_name, suffix))
            try:
                new_view.Name = target_name
            except Exception as exc:
                errors.append("Failed to rename view '{}': {}".format(view.Name, exc))

            try:
                new_view.ViewTemplateId = DB.ElementId.InvalidElementId
            except Exception as exc:
                errors.append("Failed to remove template from '{}': {}".format(target_name, exc))

            _set_view_subcategory(new_view, suffix, errors)
            duplicated.append(new_view)
    finally:
        if transaction.HasStarted():
            transaction.Commit()

    if skipped:
        errors.append("Skipped {} non-view selections.".format(len(skipped)))

    if errors:
        ui.uiUtils_alert("\n".join(errors), title=TITLE)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui.uiUtils_alert(traceback.format_exc(), title=TITLE)
