import os
import sys

from pyrevit import revit
from Autodesk.Revit import UI
from WWP_settings import get_tool_settings

script_dir = os.path.dirname(__file__)
parent_dir = os.path.dirname(script_dir)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)

from copyparameter_core import (
    TITLE,
    build_defaults,
    get_selected_elements,
    persist_defaults,
    process_elements,
    show_input_dialog,
    show_results,
)


doc = revit.doc
tool_config, save_tool_config = get_tool_settings("CopyParameter", doc=doc)


def main():
    elements = get_selected_elements()
    if not elements:
        UI.TaskDialog.Show(TITLE, "No elements selected. Please select elements and try again.")
        return

    config = show_input_dialog(
        elements,
        defaults=build_defaults(tool_config),
        title="Copy and Transform Parameter From Selected",
        empty_message="No parameters found on selected elements.",
    )
    if not config:
        return

    persist_defaults(tool_config, save_tool_config, config)
    success_count, error_count = process_elements(
        doc,
        elements,
        config,
        transaction_name="Copy and Transform Parameter From Selected",
    )
    show_results(success_count, error_count)


if __name__ == "__main__":
    main()
