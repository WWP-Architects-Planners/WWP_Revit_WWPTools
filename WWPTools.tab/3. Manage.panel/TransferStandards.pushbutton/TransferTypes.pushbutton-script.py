#! python

import os
import sys
import traceback

from System.Collections.Generic import List

from pyrevit import DB, revit
from Autodesk.Revit.DB import Architecture as DBArch


TITLE = "Transfer Standards"
HEIGHT_MULTIPLIER = 1.5


class UseDestinationTypesHandler(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


def scaled_height(value):
    try:
        return int(round(float(value) * HEIGHT_MULTIPLIER))
    except Exception:
        return value


def add_lib_path():
    script_dir = os.path.dirname(__file__)
    lib_path = None
    current = os.path.abspath(script_dir)
    for _ in range(8):
        candidate = os.path.join(current, "lib")
        if os.path.isdir(candidate):
            lib_path = candidate
            break
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    if lib_path is None:
        lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "..", "lib"))
    if lib_path not in sys.path:
        sys.path.append(lib_path)


def load_uiutils():
    add_lib_path()
    import WWP_uiUtils as ui
    return ui


def get_category_definitions():
    definitions = [
        {
            "label": "Wall Types",
            "short_label": "Walls",
            "collector": collect_element_types_by_class,
            "revit_class": DB.WallType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Curtain Wall Types",
            "short_label": "Curtain Walls",
            "collector": collect_curtain_wall_types,
            "revit_class": DB.WallType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Ceiling Types",
            "short_label": "Ceilings",
            "collector": collect_element_types_by_class,
            "revit_class": DB.CeilingType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Floor Types",
            "short_label": "Floors",
            "collector": collect_element_types_by_class,
            "revit_class": DB.FloorType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Roof Types",
            "short_label": "Roofs",
            "collector": collect_element_types_by_class,
            "revit_class": DB.RoofType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Curtain System Types",
            "short_label": "Curtain Systems",
            "collector": collect_element_types_by_class,
            "revit_class": DB.CurtainSystemType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Materials",
            "short_label": "Materials",
            "collector": collect_elements_by_class,
            "revit_class": DB.Material,
            "display_name": get_material_display_name,
            "comparison_key": get_material_comparison_key,
        },
        {
            "label": "Project Parameters",
            "short_label": "Project Parameters",
            "collector": collect_project_parameters,
            "display_name": get_project_parameter_display_name,
            "comparison_key": get_project_parameter_comparison_key,
            "copy_func": copy_project_parameter,
        },
        {
            "label": "Color Fill Schemes",
            "short_label": "Color Fill Schemes",
            "collector": collect_elements_by_class,
            "revit_class": DB.ColorFillScheme,
            "display_name": get_color_scheme_display_name,
            "comparison_key": get_color_scheme_comparison_key,
        },
        {
            "label": "View Templates",
            "short_label": "View Templates",
            "collector": collect_view_templates,
            "display_name": get_view_template_display_name,
            "comparison_key": get_view_template_comparison_key,
        },
        {
            "label": "Text Note Types",
            "short_label": "Text Note Types",
            "collector": collect_element_types_by_class,
            "revit_class": DB.TextNoteType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Dimension Styles",
            "short_label": "Dimension Styles",
            "collector": collect_element_types_by_class,
            "revit_class": DB.DimensionType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Level Types",
            "short_label": "Level Types",
            "collector": collect_element_types_by_class,
            "revit_class": DB.LevelType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Stair Types",
            "short_label": "Stair Types",
            "collector": collect_element_types_by_class,
            "revit_class": DBArch.StairsType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Handrail Types",
            "short_label": "Handrail Types",
            "collector": collect_element_types_by_class,
            "revit_class": DBArch.HandRailType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Top Rail Types",
            "short_label": "Top Rail Types",
            "collector": collect_element_types_by_class,
            "revit_class": DBArch.TopRailType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Elevation Tag Types",
            "short_label": "Elevation Tag Types",
            "collector": collect_elevation_view_family_types,
            "revit_class": DB.ViewFamilyType,
            "display_name": get_view_template_display_name,
            "comparison_key": get_view_template_comparison_key,
        },
        {
            "label": "Filled Region Types",
            "short_label": "Filled Region Types",
            "collector": collect_element_types_by_class,
            "revit_class": DB.FilledRegionType,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Fill Patterns",
            "short_label": "Fill Patterns",
            "collector": collect_elements_by_class,
            "revit_class": DB.FillPatternElement,
            "display_name": get_default_display_name,
            "comparison_key": get_default_comparison_key,
        },
        {
            "label": "Line Styles",
            "short_label": "Line Styles",
            "collector": collect_line_styles,
            "display_name": get_line_style_display_name,
            "comparison_key": get_line_style_comparison_key,
        },
    ]
    definitions.sort(key=lambda item: (item.get("label") or "").lower())
    return definitions


def iter_open_project_docs(current_doc):
    app = current_doc.Application
    for open_doc in app.Documents:
        try:
            if open_doc.IsFamilyDocument:
                continue
        except Exception:
            continue
        try:
            if open_doc.IsLinked:
                continue
        except Exception:
            pass
        if open_doc.Equals(current_doc):
            continue
        yield open_doc


def get_doc_label(doc):
    path = ""
    try:
        path = doc.PathName or ""
    except Exception:
        path = ""
    if path:
        return "{} | {}".format(doc.Title, path)
    return "{} | Unsaved Project".format(doc.Title)


def get_type_name(element):
    if element is None:
        return ""
    try:
        return DB.Element.Name.GetValue(element) or ""
    except Exception:
        pass
    try:
        return element.Name or ""
    except Exception:
        pass
    try:
        param = element.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if param:
            return param.AsString() or ""
    except Exception:
        pass
    return ""


def get_default_display_name(doc, element):
    type_name = get_type_name(element)
    family_name = ""
    try:
        family_name = element.FamilyName or ""
    except Exception:
        family_name = ""
    if family_name and family_name != type_name:
        return "{} : {}".format(family_name, type_name)
    return type_name


def normalize_name(value):
    return (value or "").strip().lower()


def get_default_comparison_key(doc, element):
    return normalize_name(get_default_display_name(doc, element))


def get_material_display_name(doc, element):
    return get_type_name(element)


def get_material_comparison_key(doc, element):
    return normalize_name(get_material_display_name(doc, element))


def get_project_parameter_binding_records(doc):
    records = []
    iterator = doc.ParameterBindings.ForwardIterator()
    while iterator.MoveNext():
        definition = iterator.Key
        binding = iterator.Current

        try:
            element = doc.GetElement(definition.Id)
        except Exception:
            element = None
        if element is None:
            continue

        try:
            is_instance = isinstance(binding, DB.InstanceBinding)
        except Exception:
            is_instance = False

        categories = []
        try:
            categories = sorted(
                [category.Name for category in binding.Categories if category and category.Name],
                key=lambda name: name.lower(),
            )
        except Exception:
            categories = []

        guid_value = ""
        try:
            guid_value = str(element.GuidValue)
        except Exception:
            guid_value = ""

        records.append(
            {
                "element": element,
                "definition": definition,
                "binding": binding,
                "name": getattr(definition, "Name", None) or get_type_name(element) or "",
                "is_instance": is_instance,
                "categories": categories,
                "guid": guid_value,
            }
        )
    return records


def get_project_parameter_display_name(doc, record):
    name = record.get("name") or "Unnamed Project Parameter"
    binding_kind = "Instance" if record.get("is_instance") else "Type"
    categories = record.get("categories") or []
    if categories:
        preview = ", ".join(categories[:3])
        if len(categories) > 3:
            preview += ", +{}".format(len(categories) - 3)
        return "{} | {} | {}".format(name, binding_kind, preview)
    return "{} | {}".format(name, binding_kind)


def get_project_parameter_comparison_key(doc, record):
    guid_value = record.get("guid") or ""
    if guid_value:
        return "guid:{}".format(guid_value.lower())
    name = normalize_name(record.get("name"))
    binding_kind = "instance" if record.get("is_instance") else "type"
    return "{}|{}".format(name, binding_kind)


def get_category_name(doc, category_id):
    if category_id is None:
        return "Unknown Category"
    try:
        category = doc.GetElement(category_id)
        if category and getattr(category, "Name", ""):
            return category.Name
    except Exception:
        pass
    try:
        return DB.LabelUtils.GetLabelFor(category_id)
    except Exception:
        return "Unknown Category"


def get_color_scheme_area_scheme_name(doc, scheme):
    try:
        area_scheme_id = getattr(scheme, "AreaSchemeId", None)
        if area_scheme_id and area_scheme_id != DB.ElementId.InvalidElementId:
            area_scheme = doc.GetElement(area_scheme_id)
            if area_scheme:
                return getattr(area_scheme, "Name", "") or ""
    except Exception:
        pass
    try:
        getter = getattr(scheme, "GetAreaSchemeId", None)
        if callable(getter):
            area_scheme_id = getter()
            if area_scheme_id and area_scheme_id != DB.ElementId.InvalidElementId:
                area_scheme = doc.GetElement(area_scheme_id)
                if area_scheme:
                    return getattr(area_scheme, "Name", "") or ""
    except Exception:
        pass
    return ""


def get_color_scheme_display_name(doc, element):
    scheme_name = get_type_name(element) or "Color Scheme"
    area_name = get_color_scheme_area_scheme_name(doc, element)
    if area_name:
        return "Area({}) : {}".format(area_name, scheme_name)
    return "{} : {}".format(get_category_name(doc, getattr(element, "CategoryId", None)), scheme_name)


def get_color_scheme_comparison_key(doc, element):
    return normalize_name(get_color_scheme_display_name(doc, element))


def get_view_template_display_name(doc, element):
    view_name = get_type_name(element)
    try:
        view_type_name = str(element.ViewType)
    except Exception:
        view_type_name = "View"
    return "{} : {}".format(view_type_name, view_name)


def get_view_template_comparison_key(doc, element):
    return normalize_name(get_view_template_display_name(doc, element))


def get_line_style_display_name(doc, element):
    try:
        category = element.GraphicsStyleCategory
        if category:
            return category.Name or get_type_name(element)
    except Exception:
        pass
    return get_type_name(element)


def get_line_style_comparison_key(doc, element):
    return normalize_name(get_line_style_display_name(doc, element))


def collect_element_types_by_class(doc, category_def):
    collector = (
        DB.FilteredElementCollector(doc)
        .OfClass(category_def["revit_class"])
        .WhereElementIsElementType()
    )
    elements = list(collector)
    elements.sort(key=lambda item: category_def["display_name"](doc, item).lower())
    return elements


def collect_elements_by_class(doc, category_def):
    collector = DB.FilteredElementCollector(doc).OfClass(category_def["revit_class"])
    elements = list(collector)
    elements.sort(key=lambda item: category_def["display_name"](doc, item).lower())
    return elements


def collect_project_parameters(doc, category_def):
    records = get_project_parameter_binding_records(doc)
    records.sort(key=lambda item: category_def["display_name"](doc, item).lower())
    return records


def collect_curtain_wall_types(doc, category_def):
    elements = collect_element_types_by_class(doc, category_def)
    filtered = []
    for item in elements:
        try:
            if item.Kind == DB.WallKind.Curtain:
                filtered.append(item)
        except Exception:
            pass
    return filtered


def collect_view_templates(doc, category_def):
    collector = DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType()
    elements = []
    for view in collector:
        try:
            if view.IsTemplate:
                elements.append(view)
        except Exception:
            pass
    elements.sort(key=lambda item: category_def["display_name"](doc, item).lower())
    return elements


def collect_elevation_view_family_types(doc, category_def):
    elements = collect_element_types_by_class(doc, category_def)
    filtered = []
    for item in elements:
        try:
            if item.ViewFamily == DB.ViewFamily.Elevation:
                filtered.append(item)
        except Exception:
            pass
    return filtered


def collect_line_styles(doc, category_def):
    elements = []
    try:
        lines_category = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
    except Exception:
        lines_category = None
    if lines_category is None:
        return elements

    try:
        subcategories = list(lines_category.SubCategories)
    except Exception:
        subcategories = []
    for subcategory in subcategories:
        try:
            graphics_style = subcategory.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        except Exception:
            graphics_style = None
        if graphics_style is not None:
            elements.append(graphics_style)
    elements.sort(key=lambda item: category_def["display_name"](doc, item).lower())
    return elements


def choose_source_document(ui, current_doc):
    source_docs = list(iter_open_project_docs(current_doc))
    if not source_docs:
        ui.uiUtils_alert(
            "Open at least one other project file in this Revit session first.",
            title=TITLE,
        )
        return None

    labels = [get_doc_label(doc) for doc in source_docs]
    selected = ui.uiUtils_select_indices(
        labels,
        title=TITLE,
        prompt="Copy types from which open project?",
        multiselect=False,
        width=920,
        height=scaled_height(420),
    )
    if not selected:
        return None
    return source_docs[selected[0]]


def choose_category(ui, current_doc, source_doc):
    category_defs = get_category_definitions()
    labels = [item["label"] for item in category_defs]
    prompt = "Target: {}\nSource: {}\n\nSelect one category to transfer:".format(
        current_doc.Title,
        source_doc.Title,
    )
    selected = ui.uiUtils_select_indices(
        labels,
        title=TITLE,
        prompt=prompt,
        multiselect=False,
        width=520,
        height=scaled_height(360),
    )
    if not selected:
        return None
    return category_defs[selected[0]]


def choose_type_elements(ui, source_doc, category_def):
    source_elements = category_def["collector"](source_doc, category_def)
    if not source_elements:
        ui.uiUtils_alert(
            "No {} found in '{}'.".format(category_def["label"].lower(), source_doc.Title),
            title=TITLE,
        )
        return []

    display_items = [category_def["display_name"](source_doc, item) for item in source_elements]
    prompt = "Select {} to copy from '{}'.".format(
        category_def["short_label"].lower(),
        source_doc.Title,
    )
    selected = ui.uiUtils_select_indices(
        display_items,
        title=TITLE,
        prompt=prompt,
        multiselect=True,
        width=980,
        height=scaled_height(680),
    )
    if not selected:
        return []
    return [source_elements[index] for index in selected]


def partition_selected_types(target_doc, selected_types, category_def):
    existing_types = category_def["collector"](target_doc, category_def)
    existing_names = {category_def["comparison_key"](target_doc, item) for item in existing_types}

    copy_candidates = []
    already_existing = []
    for source_type in selected_types:
        type_name = category_def["display_name"](target_doc, source_type)
        if category_def["comparison_key"](target_doc, source_type) in existing_names:
            already_existing.append(type_name)
        else:
            copy_candidates.append(source_type)
    return copy_candidates, sorted(set(already_existing), key=lambda name: name.lower())


def copy_type(source_doc, target_doc, source_type):
    element_ids = List[DB.ElementId]()
    element_ids.Add(source_type.Id)

    options = DB.CopyPasteOptions()
    options.SetDuplicateTypeNamesHandler(UseDestinationTypesHandler())

    copied_ids = DB.ElementTransformUtils.CopyElements(
        source_doc,
        element_ids,
        target_doc,
        DB.Transform.Identity,
        options,
    )
    return list(copied_ids) if copied_ids is not None else []


def copy_project_parameter(source_doc, target_doc, source_record):
    source_element = source_record.get("element")
    if source_element is None:
        raise Exception("Project parameter element was not found in the source document.")
    return copy_type(source_doc, target_doc, source_element)


def build_summary(category_def, source_doc, copied_names, skipped_existing, errors):
    lines = [
        "Category: {}".format(category_def["label"]),
        "Source: {}".format(source_doc.Title),
        "",
        "Copied: {}".format(len(copied_names)),
        "Already in target: {}".format(len(skipped_existing)),
        "Errors: {}".format(len(errors)),
    ]

    if copied_names:
        lines.append("")
        lines.append("Copied items:")
        lines.extend(copied_names[:15])
        if len(copied_names) > 15:
            lines.append("... ({}) more".format(len(copied_names) - 15))

    if skipped_existing:
        lines.append("")
        lines.append("Skipped because the item already exists in the target:")
        lines.extend(skipped_existing[:15])
        if len(skipped_existing) > 15:
            lines.append("... ({}) more".format(len(skipped_existing) - 15))

    if errors:
        lines.append("")
        lines.append("Errors:")
        lines.extend(errors[:12])
        if len(errors) > 12:
            lines.append("... ({}) more".format(len(errors) - 12))

    return "\n".join(lines)


def main():
    ui = load_uiutils()
    current_doc = revit.doc
    if current_doc is None:
        ui.uiUtils_alert("No active Revit document found.", title=TITLE)
        return

    source_doc = choose_source_document(ui, current_doc)
    if source_doc is None:
        return

    category_def = choose_category(ui, current_doc, source_doc)
    if category_def is None:
        return

    selected_types = choose_type_elements(ui, source_doc, category_def)
    if not selected_types:
        return

    copy_candidates, skipped_existing = partition_selected_types(current_doc, selected_types, category_def)
    if skipped_existing:
        message = [
            "{} selected type(s) already exist in '{}'.".format(len(skipped_existing), current_doc.Title),
            "",
            "Their existing destination versions will be kept.",
            "",
            "Continue copying the remaining {} type(s)?".format(len(copy_candidates)),
        ]
        if not ui.uiUtils_confirm("\n".join(message), title=TITLE):
            return

    if not copy_candidates:
        ui.uiUtils_alert(
            build_summary(category_def, source_doc, [], skipped_existing, []),
            title=TITLE,
        )
        return

    copied_names = []
    errors = []
    transaction_name = "Transfer {}".format(category_def["label"])
    copy_func = category_def.get("copy_func", copy_type)
    with revit.Transaction(transaction_name):
        for source_type in copy_candidates:
            type_name = category_def["display_name"](source_doc, source_type) or "<Unnamed Item>"
            try:
                copy_func(source_doc, current_doc, source_type)
                copied_names.append(type_name)
            except Exception as exc:
                errors.append("{}: {}".format(type_name, str(exc)))

    ui.uiUtils_alert(
        build_summary(category_def, source_doc, copied_names, skipped_existing, errors),
        title=TITLE,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception:
        ui = None
        try:
            ui = load_uiutils()
        except Exception:
            ui = None
        if ui is not None:
            ui.uiUtils_alert(traceback.format_exc(), title=TITLE)
        else:
            print(traceback.format_exc())
