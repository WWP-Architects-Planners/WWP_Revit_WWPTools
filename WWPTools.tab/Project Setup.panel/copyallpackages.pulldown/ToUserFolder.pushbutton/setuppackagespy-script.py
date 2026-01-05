import os
import shutil
import ctypes

try:
    from pyrevit import HOST_APP
    exe_year = int(HOST_APP.version)
except Exception:
    try:
        exe_year = int(__revit__.Application.VersionNumber)
    except Exception:
        exe_year = 2024


def _parse_version_tuple(version_str):
    parts = (version_str or "").split(".")
    out = []
    for part in parts:
        try:
            out.append(int(part))
        except Exception:
            out.append(-1)
    return tuple(out)


def _find_latest_dynamo_packages_dir(dest_base):
    dynamo_root = os.path.join(dest_base, r"Dynamo\Dynamo Revit")
    if not os.path.isdir(dynamo_root):
        return None
    versions = [
        d for d in os.listdir(dynamo_root)
        if os.path.isdir(os.path.join(dynamo_root, d))
    ]
    if not versions:
        return None
    latest = max(versions, key=_parse_version_tuple)
    return os.path.join(dynamo_root, latest, "packages")

def copy_files(source, destination):
    for root, dirs, files in os.walk(source):
        relative_path = os.path.relpath(root, source)
        dest_folder = os.path.join(destination, relative_path)
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
        for file in files:
            src_file = os.path.join(root, file)
            dest_file = os.path.join(dest_folder, file)
            if not os.path.exists(dest_file):
                shutil.copy2(src_file, dest_file)

# Show a confirmation message box
confirm_response = ctypes.windll.user32.MessageBoxW(None,
                                                   "It will start the process of copying Dynamo packages. \nMake sure you run dynamo application at least one time.\nAre you sure to run the copy process now?",
                                                   "Confirmation",
                                                   1 | 0x30)  # 1: OK, 0x30: Information icon

# Convert the response code to 'y' or 'n'
if confirm_response == 1:
    source_base = r"N:\Library\Design Software\Autodesk\Revit\Dynamo\packages"
    source_all_version = os.path.join(source_base, "All Version")
    dest_base = os.path.expandvars("%appdata%")

    if exe_year < 2024:
        print("This tool is intended for Revit 2024+ only.")
    else:
        source_dir = os.path.join(source_base, "Revit {}".format(exe_year))
        dest_dir = _find_latest_dynamo_packages_dir(dest_base)

        if not os.path.isdir(source_dir):
            print("Source packages folder not found: \n{}\n".format(source_dir))
        elif not dest_dir:
            print("Could not find Dynamo packages folder under AppData.\n")
        else:
            print("Copying files from: \n{}\n".format(source_dir))
            print("To destination: \n{}\n".format(dest_dir))

            if os.path.exists(dest_dir):
                print("Deleting existing files in: \n{}\n".format(dest_dir))
                shutil.rmtree(dest_dir, ignore_errors=True)
            os.makedirs(dest_dir, exist_ok=True)
            copy_files(source_dir, dest_dir)
            print("Files copied successfully.\n")

            print("Copying extra folder from: \n{}\n".format(source_all_version))
            print("To destination: \n{}\n".format(dest_dir))
            copy_files(source_all_version, dest_dir)
            print("Folder copied\nsuccessfully.\n")
else:
    print("Copy process cancelled.")
print("\nAll Process is done")
