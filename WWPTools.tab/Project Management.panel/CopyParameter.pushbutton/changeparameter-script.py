"""
Copy and transform instance parameters with find/replace functionality.
Copies a source parameter value to a target parameter with optional text replacement.
"""

from pyrevit import revit, DB, forms, script
from typing import Optional, List

# Get the active document
doc = revit.doc
output = script.get_output()

def get_selected_elements() -> List[DB.Element]:
    """Get currently selected elements from the Revit model."""
    try:
        selection = revit.get_selection()
        return list(selection.elements)
    except Exception as e:
        forms.alert(f"Error getting selection: {str(e)}")
        return []

def get_parameter_value(element: DB.Element, param_name: str) -> Optional[str]:
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

def set_parameter_value(element: DB.Element, param_name: str, value: str) -> bool:
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

def show_input_dialog() -> Optional[dict]:
    """
    Display user input dialog for parameter names and find/replace values.
    
    Returns:
        Dictionary with 'source_param', 'target_param', 'find_text', 'replace_text'
        or None if cancelled
    """
    from pyrevit.forms import WPFWindow
    
    # Create simple input dialog
    components = [
        ("Source Parameter", "View Name"),
        ("Target Parameter", "Title on Sheet"),
        ("Find Text", ""),
        ("Replace Text", "")
    ]
    
    # Use TaskDialog for simpler input
    dialog = DB.TaskDialog("Copy and Transform Parameter")
    dialog.MainInstruction = "Enter the parameters and find/replace values"
    dialog.MainContent = "Please configure the parameter copy operation"
    
    # For now, use a workaround with multiple dialogs
    source_param = forms.ask_for_one_item(
        ["View Name", "Title on Sheet"],
        "Select Source Parameter:"
    )
    
    if not source_param:
        return None
    
    target_param = forms.ask_for_one_item(
        ["View Name", "Title on Sheet"],
        "Select Target Parameter:"
    )
    
    if not target_param:
        return None
    
    find_text = forms.ask_for_string(
        default="",
        prompt="Text to find (leave empty for no replacement):",
        title="Find Text"
    )
    
    replace_text = forms.ask_for_string(
        default="",
        prompt="Text to replace with:",
        title="Replace Text"
    )
    
    return {
        "source_param": source_param,
        "target_param": target_param,
        "find_text": find_text or "",
        "replace_text": replace_text or ""
    }

def main():
    """Main execution function."""
    # Get selected elements
    elements = get_selected_elements()
    
    if not elements:
        forms.alert("No elements selected. Please select elements and try again.")
        return
    
    # Get user input
    config = show_input_dialog()
    if not config:
        return
    
    source_param = config["source_param"]
    target_param = config["target_param"]
    find_text = config["find_text"]
    replace_text = config["replace_text"]
    
    # Process elements
    success_count = 0
    error_count = 0
    
    with DB.Transaction(doc, "Copy and Transform Parameter") as txn:
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
            
            # Set target parameter
            if set_parameter_value(element, target_param, target_value):
                print(f"Element {element.Id}: '{source_param}' → '{target_param}' = '{target_value}'")
                success_count += 1
            else:
                print(f"Element {element.Id}: Failed to set parameter '{target_param}'")
                error_count += 1
        
        txn.Commit()
    
    # Show results
    message = f"Operation Complete:\n\n✓ Successful: {success_count}\n✗ Failed: {error_count}"
    forms.alert(message, title="Parameter Copy Results")
    
    print("\n" + "="*50)
    print(message)
    print("="*50)

if __name__ == "__main__":
    main()
