#! python3
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import Autodesk.Revit.DB as DB
from Autodesk.Revit.UI import TaskDialog
import os
import datetime
import traceback
import WWP_uiUtils as ui

doc = __revit__.ActiveUIDocument.Document
app = doc.Application

# ============================================================
# Configuration
# ============================================================
DEFAULT_SHARED_PARAMETERS_PATH = r"N:\Library\Design Software\Autodesk\Revit\Shared Parameters\SharedParameters.txt"
TEMPLATE_FILE_PATH = r"N:\Library\Design Software\Autodesk\Revit\Standards\Project Parameters\Project Parameters Template.xlsx"

# ============================================================
# File Dialog Helpers
# ============================================================
def _pick_first_existing_path(paths):
    for p in paths:
        try:
            if p and os.path.exists(p):
                return p
        except Exception:
            continue
    return None


def _choose_shared_parameter_file(default_path):
    initial_dir = ""
    try:
        if default_path:
            initial_dir = os.path.dirname(default_path)
    except Exception:
        pass
    return ui.uiUtils_open_file_dialog(
        title="Select Shared Parameters File",
        filter_text="Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
        multiselect=False,
        initial_directory=initial_dir or "",
    )


def _choose_save_excel_file():
    default_dir = r"N:\Library\Design Software\Autodesk\Revit\Shared Parameters"
    default_filename = "Project_Parameters_Template_{}.xlsx".format(
        datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    )
    return ui.uiUtils_save_file_dialog(
        title="Save Project Parameter Template",
        filter_text="Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
        default_extension="xlsx",
        initial_directory=default_dir if os.path.isdir(default_dir) else "",
        file_name=default_filename,
    )


# ============================================================
# Shared Parameters Reader
# ============================================================
def _read_shared_parameters(file_path):
    """Read shared parameters file and extract parameters with their groups."""
    parameters = []
    groups = {}  # Store group ID -> group name mapping
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        print("Read file with utf-8 encoding, {} lines".format(len(lines)))
    except Exception as e1:
        print("utf-8 failed: {}".format(str(e1)))
        try:
            with open(file_path, 'r', encoding='utf-16') as f:
                lines = f.readlines()
            print("Read file with utf-16 encoding, {} lines".format(len(lines)))
        except Exception as e2:
            print("utf-16 failed: {}".format(str(e2)))
            with open(file_path, 'r') as f:
                lines = f.readlines()
            print("Read file with default encoding, {} lines".format(len(lines)))
    
    in_groups = False
    in_params = False
    
    for line_num, line in enumerate(lines, start=1):
        line = line.strip()
        
        if not line:
            continue
            
        # Check for section headers (lines starting with *)
        if line.startswith('*GROUP'):
            in_groups = True
            in_params = False
            print("Entering GROUP section at line {}".format(line_num))
            continue
            
        if line.startswith('*PARAM'):
            in_groups = False
            in_params = True
            print("Entering PARAM section at line {}, found {} groups".format(line_num, len(groups)))
            continue
        
        # Skip other metadata lines
        if line.startswith('*'):
            continue
        
        # Process GROUP data lines
        if in_groups and line.startswith('GROUP'):
            parts = line.split('\t')
            if len(parts) >= 3:
                # Format: GROUP ID NAME
                group_id = parts[1]
                group_name = parts[2]
                groups[group_id] = group_name
                print("  Group {}: {}".format(group_id, group_name))
        
        # Process PARAM data lines
        if in_params and line.startswith('PARAM'):
            parts = line.split('\t')
            if len(parts) >= 6:
                # Format: PARAM GUID NAME DATATYPE DATACATEGORY GROUP ...
                param_guid = parts[1]
                param_name = parts[2]
                param_type = parts[3]
                data_category = parts[4]
                group_id = parts[5]
                
                # Resolve group name from stored mapping
                group_name = groups.get(group_id, "")
                
                parameters.append({
                    'guid': param_guid,
                    'name': param_name,
                    'type': param_type,
                    'data_category': data_category,
                    'group': group_name
                })
    
    print("Parsed {} parameters from {} groups".format(len(parameters), len(groups)))
    return parameters


    print("Parsed {} parameters from {} groups".format(len(parameters), len(groups)))
    return parameters


# ============================================================
# Excel Writer (COM)
# ============================================================
def _populate_template(template_path, output_path, parameters):
    """Copy template, clear Variables data, and populate with shared parameters."""
    import shutil
    
    try:
        clr.AddReference('System')
        from System import Activator, Type
        from System.Runtime.InteropServices import Marshal
        Reflection = __import__('System.Reflection', fromlist=['BindingFlags'])
        BindingFlags = Reflection.BindingFlags
    except Exception as e:
        raise Exception("Excel COM not available. ({})".format(str(e)))

    _BASE_FLAGS = BindingFlags.Public | BindingFlags.Instance | BindingFlags.OptionalParamBinding

    def _com_get(obj, name):
        try:
            return obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.GetProperty, None, obj, None)
        except Exception:
            return obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.InvokeMethod, None, obj, [])

    def _com_set(obj, name, value):
        try:
            obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.SetProperty, None, obj, [value])
        except Exception:
            obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.InvokeMethod, None, obj, [value])

    def _com_call(obj, name, *args):
        try:
            return obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.InvokeMethod, None, obj, list(args))
        except Exception:
            return obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.GetProperty, None, obj, list(args))

    excel_app = None
    workbook = None
    ws_variables = None

    try:
        # First, copy the template file
        print("Copying template file...")
        shutil.copy2(template_path, output_path)
        print("Template copied to: {}".format(output_path))
        
        excel_type = Type.GetTypeFromProgID('Excel.Application')
        if excel_type is None:
            raise Exception("Excel COM ProgID not found (is Microsoft Excel installed?)")
        excel_app = Activator.CreateInstance(excel_type)
        _com_set(excel_app, 'Visible', False)
        _com_set(excel_app, 'DisplayAlerts', False)

        # Open the copied file
        print("Opening copied template...")
        workbooks = _com_get(excel_app, 'Workbooks')
        workbook = _com_call(workbooks, 'Open', output_path)
        worksheets = _com_get(workbook, 'Worksheets')
        
        # Get Variables sheet
        try:
            ws_variables = _com_call(worksheets, 'Item', 'Variables')
            print("Variables sheet found")
        except Exception:
            raise Exception("Variables sheet not found in template")
        
        # Clear existing data (keep headers in row 1, preserve columns I-N)
        print("Clearing existing variable data (columns A-H only)...")
        used_range = _com_get(ws_variables, 'UsedRange')
        row_count = int(_com_get(_com_get(used_range, 'Rows'), 'Count'))
        if row_count > 1:
            clear_range = _com_call(ws_variables, 'Range', "A2:H{}".format(row_count))
            _com_call(clear_range, 'ClearContents')
        
        # Write parameter data - mapping to template column structure
        print("Writing {} parameters to Variables sheet...".format(len(parameters)))
        for row_idx, param in enumerate(parameters, start=2):
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 1), 'Value2', param['name'])              # Parameter Name
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 2), 'Value2', param['group'])             # ParameterGroup
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 3), 'Value2', param['type'])              # ParameterType
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 4), 'Value2', param['type'])              # ParameterType Revit
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 5), 'Value2', param['data_category'])     # DataType
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 6), 'Value2', param['data_category'])     # DataType Revit
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 7), 'Value2', param['type'])              # ParameterType
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 8), 'Value2', param['type'])              # ParameterType Revit
        
        # Save using Ctrl+S equivalent (avoids SaveAs which triggers O365 label prompt)
        print("Saving file...")
        try:
            _com_call(workbook, 'Save')
            print("Save completed")
        except Exception as save_ex:
            print("Save failed: {}".format(str(save_ex)))
            raise
        
        # Close workbook
        try:
            if workbook is not None:
                print("Closing workbook...")
                _com_call(workbook, 'Close', True)
                print("Workbook closed")
        except Exception as close_ex:
            print("Close failed: {}".format(str(close_ex)))
        
        # Quit Excel
        try:
            if excel_app is not None:
                print("Quitting Excel...")
                _com_call(excel_app, 'Quit')
                print("Excel quit")
        except Exception as quit_ex:
            print("Quit failed: {}".format(str(quit_ex)))
        
        # Verify file
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print("File saved successfully: {} bytes".format(file_size))
            return output_path
        else:
            raise Exception("File not found after save")

    except Exception as ex:
        print("Exception: {}".format(str(ex)))
        raise
    finally:
        try:
            if ws_variables is not None:
                Marshal.ReleaseComObject(ws_variables)
        except Exception:
            pass
        try:
            if workbook is not None:
                Marshal.ReleaseComObject(workbook)
        except Exception:
            pass
        try:
            if excel_app is not None:
                Marshal.ReleaseComObject(excel_app)
        except Exception:
            pass
    """Create Excel template with Variables sheet and Project Parameters sheet."""
    try:
        clr.AddReference('System')
        from System import Activator, Type
        from System.Runtime.InteropServices import Marshal
        Reflection = __import__('System.Reflection', fromlist=['BindingFlags'])
        BindingFlags = Reflection.BindingFlags
    except Exception as e:
        raise Exception("Excel COM not available. ({})".format(str(e)))

    _BASE_FLAGS = BindingFlags.Public | BindingFlags.Instance | BindingFlags.OptionalParamBinding

    def _com_get(obj, name):
        try:
            return obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.GetProperty, None, obj, None)
        except Exception:
            return obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.InvokeMethod, None, obj, [])

    def _com_set(obj, name, value):
        try:
            obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.SetProperty, None, obj, [value])
        except Exception:
            obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.InvokeMethod, None, obj, [value])

    def _com_call(obj, name, *args):
        try:
            return obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.InvokeMethod, None, obj, list(args))
        except Exception:
            return obj.GetType().InvokeMember(name, _BASE_FLAGS | BindingFlags.GetProperty, None, obj, list(args))

    excel_app = None
    workbook = None
    ws_variables = None
    ws_project = None

    try:
        excel_type = Type.GetTypeFromProgID('Excel.Application')
        if excel_type is None:
            raise Exception("Excel COM ProgID not found (is Microsoft Excel installed?)")
        excel_app = Activator.CreateInstance(excel_type)
        _com_set(excel_app, 'Visible', False)
        _com_set(excel_app, 'DisplayAlerts', False)

        workbooks = _com_get(excel_app, 'Workbooks')
        workbook = _com_call(workbooks, 'Add')

        worksheets = _com_get(workbook, 'Worksheets')
        
        # Create Variables sheet
        ws_variables = _com_call(worksheets, 'Item', 1)
        _com_set(ws_variables, 'Name', 'Variables')
        
        # Write Variables headers
        variables_headers = ['Parameter Name', 'Group Name', 'Parameter Type', 'Data Category', 'GUID']
        for col_idx, header in enumerate(variables_headers, start=1):
            cell = _com_call(ws_variables, 'Cells', 1, col_idx)
            _com_set(cell, 'Value2', header)
            font = _com_get(cell, 'Font')
            _com_set(font, 'Bold', True)
        
        # Write parameter data
        for row_idx, param in enumerate(parameters, start=2):
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 1), 'Value2', param['name'])
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 2), 'Value2', param['group'])
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 3), 'Value2', param['type'])
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 4), 'Value2', param['data_category'])
            _com_set(_com_call(ws_variables, 'Cells', row_idx, 5), 'Value2', param['guid'])
        
        # Auto-fit columns in Variables sheet
        cols = _com_get(ws_variables, 'Columns')
        _com_call(cols, 'AutoFit')
        
        # Create Project Parameters sheet
        ws_project = _com_call(worksheets, 'Add')
        _com_set(ws_project, 'Name', 'Project Parameters')
        
        # Write Project Parameters headers (matching the import script expectations)
        project_headers = [
            'Parameter Name',
            'Group Name',
            'Discipline',
            'ParaTypeRevit',
            'ParaTypeSharedRevit',
            'Group',
            'Instance Parameter',
            'Reporting Parameter',
            'Vary group / Type',
            'Category API'
        ]
        for col_idx, header in enumerate(project_headers, start=1):
            cell = _com_call(ws_project, 'Cells', 1, col_idx)
            _com_set(cell, 'Value2', header)
            font = _com_get(cell, 'Font')
            _com_set(font, 'Bold', True)
        
        # Add data validation and formulas for columns
        # Add 100 empty rows for user to fill
        param_count = len(parameters)
        var_range = "Variables!$A$2:$A${}".format(param_count + 1)
        
        # Column A (Parameter Name) - dropdown validation
        print("Adding data validation for Parameter Name column...")
        col_a_range = _com_call(ws_project, 'Range', "A2:A101")
        dv = _com_get(col_a_range, 'Validation')
        try:
            _com_call(dv, 'Delete')
        except Exception:
            pass
        try:
            # Type=3 is xlList, AlertStyle=2 is xlValidAlertStop
            _com_call(dv, 'Add', 3, 2, None, var_range)
            _com_set(dv, 'IgnoreBlank', True)
            _com_set(dv, 'InCellDropdown', True)
            print("Data validation added successfully")
        except Exception as dv_ex:
            print("Data validation failed: {}".format(str(dv_ex)))
        
        # Column B (Group Name) - VLOOKUP formula
        print("Adding VLOOKUP formula for Group Name column...")
        for row in range(2, 102):
            cell_b = _com_call(ws_project, 'Cells', row, 2)
            _com_set(cell_b, 'Formula', '=IFERROR(VLOOKUP(A{},Variables!$A$2:$E${},2,FALSE),"")'.format(row, param_count + 1))
        
        # Column D (ParaTypeRevit) - VLOOKUP formula for Parameter Type
        print("Adding VLOOKUP formula for ParaTypeRevit column...")
        for row in range(2, 102):
            cell_d = _com_call(ws_project, 'Cells', row, 4)
            _com_set(cell_d, 'Formula', '=IFERROR(VLOOKUP(A{},Variables!$A$2:$E${},3,FALSE),"")'.format(row, param_count + 1))
        
        # Auto-fit columns in Project Parameters sheet
        cols = _com_get(ws_project, 'Columns')
        _com_call(cols, 'AutoFit')
        
        # Make Project Parameters sheet active
        _com_call(ws_project, 'Activate')
        
        # Save and close workbook (51 = xlOpenXMLWorkbook format for .xlsx)
        print("Saving workbook to: {}".format(output_path))
        try:
            _com_call(workbook, 'SaveAs', output_path, 51)
            print("SaveAs completed")
        except Exception as save_ex:
            print("SaveAs failed: {}".format(str(save_ex)))
            # Try alternative save method
            try:
                _com_call(workbook, 'SaveAs', output_path)
                print("SaveAs without format completed")
            except Exception as save_ex2:
                print("Alternative SaveAs also failed: {}".format(str(save_ex2)))
                raise
        
        # Close workbook to flush to disk
        try:
            if workbook is not None:
                print("Closing workbook...")
                _com_call(workbook, 'Close', True)  # True = save changes
                print("Workbook closed")
        except Exception as close_ex:
            print("Close failed: {}".format(str(close_ex)))
        
        # Quit Excel
        try:
            if excel_app is not None:
                print("Quitting Excel...")
                _com_call(excel_app, 'Quit')
                print("Excel quit")
        except Exception as quit_ex:
            print("Quit failed: {}".format(str(quit_ex)))
        
        # Verify file exists after close
        if os.path.exists(output_path):
            print("File verified after close: {}".format(output_path))
            print("File size: {} bytes".format(os.path.getsize(output_path)))
        else:
            print("ERROR: File not found after close!")
        
        return output_path

    except Exception as ex:
        print("Exception in _create_excel_template: {}".format(str(ex)))
        raise
    finally:
        # Clean up COM objects
        try:
            if ws_variables is not None:
                Marshal.ReleaseComObject(ws_variables)
        except Exception:
            pass
        try:
            if ws_project is not None:
                Marshal.ReleaseComObject(ws_project)
        except Exception:
            pass
        try:
            if workbook is not None:
                Marshal.ReleaseComObject(workbook)
        except Exception:
            pass
        try:
            if excel_app is not None:
                Marshal.ReleaseComObject(excel_app)
        except Exception:
            pass


# ============================================================
# Main
# ============================================================
try:
    print("Create Project Parameter Template - Starting...")
    print("Started: {}".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    
    # Select shared parameters file
    default_shared = _pick_first_existing_path([DEFAULT_SHARED_PARAMETERS_PATH])
    selected_shared = _choose_shared_parameter_file(default_shared or DEFAULT_SHARED_PARAMETERS_PATH)
    
    if not selected_shared:
        selected_shared = default_shared
    
    if not selected_shared or not os.path.exists(selected_shared):
        TaskDialog.Show(
            "Error",
            "Shared Parameters file not found.\n\nDefault:\n{}".format(default_shared or DEFAULT_SHARED_PARAMETERS_PATH),
        )
        raise Exception("Shared Parameters file not found")
    
    print("Shared Parameters: {}".format(selected_shared))
    
    # Select output Excel file
    output_file = _choose_save_excel_file()
    if not output_file:
        TaskDialog.Show("Cancelled", "Template creation cancelled.")
        raise Exception("Template creation cancelled")
    
    print("Output File: {}".format(output_file))
    
    # Read shared parameters
    print("Reading shared parameters...")
    parameters = _read_shared_parameters(selected_shared)
    print("Found {} parameters".format(len(parameters)))
    
    # Check if template exists
    if not os.path.exists(TEMPLATE_FILE_PATH):
        TaskDialog.Show(
            "Error",
            "Template file not found:\n{}".format(TEMPLATE_FILE_PATH),
        )
        raise Exception("Template file not found")
    
    # Populate template
    print("Populating template...")
    _populate_template(TEMPLATE_FILE_PATH, output_file, parameters)
    
    print("Template created successfully!")
    print("Completed: {}".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    
    TaskDialog.Show(
        "Success",
        "Template created successfully!\n\nParameters: {}\nLocation:\n{}".format(
            len(parameters),
            output_file
        )
    )

except Exception as e:
    print("FATAL ERROR: {}".format(str(e)))
    print("Error Type: {}".format(type(e).__name__))
    try:
        print(traceback.format_exc())
    except Exception:
        pass
    TaskDialog.Show("Error", "Failed to create template:\n{}".format(str(e)))
