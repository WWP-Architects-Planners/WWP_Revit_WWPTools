#! python3
from pyrevit import revit, DB
from Autodesk.Revit import UI
import WWP_uiUtils as ui


TITLE = "Copy Parameter"


def ask_for_inputs(param_names, defaults=None, title="Copy and Transform Parameter"):
    defaults = defaults or {}
    try:
        return ui.uiUtils_parameter_copy_inputs(
            param_names=param_names or [],
            title=title,
            source_default=defaults.get("source_param") or "View Name",
            target_default=defaults.get("target_param") or "Title on Sheet",
            find_default=defaults.get("find_text") or "",
            replace_default=defaults.get("replace_text") or "",
            prefix_default=defaults.get("prefix") or "",
            suffix_default=defaults.get("suffix") or "",
            width=480,
            height=460,
        )
    except Exception as ex:
        UI.TaskDialog.Show(TITLE, "Unable to load input dialog: {}".format(str(ex)))
        return None


def get_all_parameter_names(elements):
    names = set()
    for element in elements:
        try:
            for param in element.Parameters:
                if param and param.Definition:
                    names.add(param.Definition.Name)
        except Exception:
            continue
    return sorted(names)


def get_selected_elements():
    try:
        selection = revit.get_selection()
        return list(selection.elements)
    except Exception as ex:
        UI.TaskDialog.Show(TITLE, "Error getting selection: {}".format(str(ex)))
        return []


def get_parameter_value(element, param_name):
    try:
        param = element.LookupParameter(param_name)
        if param is None:
            return None
        if param.StorageType == DB.StorageType.String:
            return str(param.AsString())
        return str(param.AsValueString())
    except Exception as ex:
        print("Error getting parameter '{}': {}".format(param_name, str(ex)))
        return None


def set_parameter_value(element, param_name, value):
    try:
        param = element.LookupParameter(param_name)
        if param is None or param.IsReadOnly:
            return False
        param.Set(value)
        return True
    except Exception as ex:
        print("Error setting parameter '{}': {}".format(param_name, str(ex)))
        return False


def show_input_dialog(elements, defaults=None, title="Copy and Transform Parameter", empty_message=None):
    param_names = get_all_parameter_names(elements)
    if not param_names:
        UI.TaskDialog.Show(TITLE, empty_message or "No parameters found.")
        return None
    return ask_for_inputs(param_names, defaults=defaults, title=title)


def build_defaults(tool_config):
    return {
        "source_param": getattr(tool_config, "source_param", "") or "View Name",
        "target_param": getattr(tool_config, "target_param", "") or "Title on Sheet",
        "find_text": getattr(tool_config, "find_text", "") or "",
        "replace_text": getattr(tool_config, "replace_text", "") or "",
        "prefix": getattr(tool_config, "prefix", "") or "",
        "suffix": getattr(tool_config, "suffix", "") or "",
    }


def persist_defaults(tool_config, save_tool_config, config):
    tool_config.source_param = config.get("source_param") or ""
    tool_config.target_param = config.get("target_param") or ""
    tool_config.find_text = config.get("find_text") or ""
    tool_config.replace_text = config.get("replace_text") or ""
    tool_config.prefix = config.get("prefix") or ""
    tool_config.suffix = config.get("suffix") or ""
    save_tool_config()


def process_elements(doc, elements, config, transaction_name):
    source_param = config["source_param"]
    target_param = config["target_param"]
    find_text = config.get("find_text") or ""
    replace_text = config.get("replace_text") or ""
    prefix = config.get("prefix") or ""
    suffix = config.get("suffix") or ""

    success_count = 0
    error_count = 0

    txn = DB.Transaction(doc, transaction_name)
    txn.Start()
    try:
        for element in elements:
            source_value = get_parameter_value(element, source_param)
            if source_value is None:
                print("Element {}: Parameter '{}' not found".format(element.Id, source_param))
                error_count += 1
                continue

            target_value = source_value.replace(find_text, replace_text) if find_text else source_value
            if prefix or suffix:
                target_value = "{}{}{}".format(prefix, target_value, suffix)

            if set_parameter_value(element, target_param, target_value):
                print("Element {}: '{}' -> '{}' = '{}'".format(
                    element.Id,
                    source_param,
                    target_param,
                    target_value,
                ))
                success_count += 1
            else:
                print("Element {}: Failed to set parameter '{}'".format(element.Id, target_param))
                error_count += 1
        txn.Commit()
    except Exception:
        txn.RollBack()
        raise

    return success_count, error_count


def show_results(success_count, error_count, title="Parameter Copy Results"):
    message = "Operation Complete:\n\nSuccessful: {}\nFailed: {}".format(success_count, error_count)
    UI.TaskDialog.Show(title, message)
    print("\n" + "=" * 50)
    print(message)
    print("=" * 50)


def collect_category_records(doc):
    categories = {}
    collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
    for element in collector:
        try:
            category = element.Category
        except Exception:
            category = None
        if category is None:
            continue
        name = (category.Name or "").strip()
        if not name:
            continue
        key = int(category.Id.IntegerValue)
        if key not in categories:
            categories[key] = {
                "id": category.Id,
                "name": name,
                "count": 0,
            }
        categories[key]["count"] += 1
    return sorted(categories.values(), key=lambda item: item["name"].lower())


def choose_category(doc, title="Copy Parameter by Category"):
    records = collect_category_records(doc)
    if not records:
        UI.TaskDialog.Show(TITLE, "No categories with elements were found in this project.")
        return None

    labels = ["{} ({})".format(record["name"], record["count"]) for record in records]
    selected = ui.uiUtils_select_indices(
        labels,
        title=title,
        prompt="Select a category to process:",
        multiselect=False,
        width=520,
        height=620,
    )
    if not selected:
        return None
    index = int(selected[0])
    if index < 0 or index >= len(records):
        return None
    return records[index]


def get_elements_by_category(doc, category_id):
    results = []
    try:
        collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
        collector = collector.WherePasses(DB.ElementCategoryFilter(category_id))
        for element in collector:
            try:
                if element.Category and element.Category.Id == category_id:
                    results.append(element)
            except Exception:
                continue
    except Exception as ex:
        UI.TaskDialog.Show(TITLE, "Failed to collect category elements: {}".format(str(ex)))
        return []
    return results
