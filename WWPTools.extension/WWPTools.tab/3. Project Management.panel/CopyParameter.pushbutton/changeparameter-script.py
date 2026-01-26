#! python3
from pyrevit import revit, DB, script
from Autodesk.Revit import UI
import WWP_uiUtils as ui
# Get the active document
doc = revit.doc
output = script.get_output()


def ask_for_inputs(param_names):
    """Single form for parameter mapping and text transforms."""
    try:
        return ui.uiUtils_parameter_copy_inputs(
            param_names=param_names or [],
            title="Copy and Transform Parameter",
            source_default="View Name",
            target_default="Title on Sheet",
            find_default="",
            replace_default="",
            prefix_default="",
            suffix_default="",
            width=480,
            height=460,
        )
    except Exception as e:
        UI.TaskDialog.Show("Copy Parameter", "Unable to load input dialog: {}".format(str(e)))
        return None

def get_all_parameter_names(elements):
    """Return a sorted list of parameter names across all selected elements."""
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
    """Get currently selected elements from the Revit model."""
    try:
        selection = revit.get_selection()
        return list(selection.elements)
    except Exception as e:
        UI.TaskDialog.Show("Copy Parameter", f"Error getting selection: {str(e)}")
        return []

def get_parameter_value(element, param_name):
    """
    Get the value of a parameter from an element.
    
    Args:
        element: The Revit element
        param_name: Name of the parameter to retrieve
        
    Returns:
        The parameter value as string, or None if not found
    """
    try:
        param = element.LookupParameter(param_name)
        if param is None:
            return None
        return str(param.AsString()) if param.StorageType == DB.StorageType.String else str(param.AsValueString())
    except Exception as e:
        print(f"Error getting parameter '{param_name}': {str(e)}")
        return None

def set_parameter_value(element, param_name, value):
    """
    Set the value of a parameter on an element.
    
    Args:
        element: The Revit element
        param_name: Name of the parameter to set
        value: The value to set
        
    Returns:
        True if successful, False otherwise
    """
    try:
        param = element.LookupParameter(param_name)
        if param is None:
            return False
        if param.IsReadOnly:
            return False
        param.Set(value)
        return True
    except Exception as e:
        print(f"Error setting parameter '{param_name}': {str(e)}")
        return False

def show_input_dialog(elements):
    """
    Display user input dialog for parameter names and find/replace values.
    
    Returns:
        Dictionary with 'source_param', 'target_param', 'find_text', 'replace_text'
        or None if cancelled
    """
    param_names = get_all_parameter_names(elements)
    if not param_names:
        UI.TaskDialog.Show("Copy Parameter", "No parameters found on selected elements.")
        return None
    return ask_for_inputs(param_names)

def main():
    """Main execution function."""
    # Get selected elements
    elements = get_selected_elements()
    
    if not elements:
        UI.TaskDialog.Show("Copy Parameter", "No elements selected. Please select elements and try again.")
        return
    
    # Get user input
    config = show_input_dialog(elements)
    if not config:
        return
    
    source_param = config["source_param"]
    target_param = config["target_param"]
    find_text = config["find_text"] or ""
    replace_text = config["replace_text"] or ""
    prefix = config.get("prefix") or ""
    suffix = config.get("suffix") or ""
    
    # Process elements
    success_count = 0
    error_count = 0
    
    txn = DB.Transaction(doc, "Copy and Transform Parameter")
    txn.Start()

    for element in elements:
        # Get source parameter value
        source_value = get_parameter_value(element, source_param)

        if source_value is None:
            print(f"Element {element.Id}: Parameter '{source_param}' not found")
            error_count += 1
            continue

        # Apply find/replace if specified
        if find_text:
            target_value = source_value.replace(find_text, replace_text)
        else:
            target_value = source_value

        if prefix or suffix:
            target_value = f"{prefix}{target_value}{suffix}"

        # Set target parameter
        if set_parameter_value(element, target_param, target_value):
            print(f"Element {element.Id}: '{source_param}' -> '{target_param}' = '{target_value}'")
            success_count += 1
        else:
            print(f"Element {element.Id}: Failed to set parameter '{target_param}'")
            error_count += 1

    txn.Commit()
    
    # Show results
    message = f"Operation Complete:\n\nSuccessful: {success_count}\nFailed: {error_count}"
    UI.TaskDialog.Show("Parameter Copy Results", message)
    
    print("\n" + "="*50)
    print(message)
    print("="*50)

if __name__ == "__main__":
    main()
