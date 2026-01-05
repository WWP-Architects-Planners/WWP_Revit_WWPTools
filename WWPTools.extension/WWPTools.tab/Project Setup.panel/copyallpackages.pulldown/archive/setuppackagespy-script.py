import os
import shutil

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

def delete_files_in_folder(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print("Failed to delete " + file_path + ". Reason: " + str(e))

overwrite = input("Do you want to overwrite existing files? (Y/N): ")

source_dir = r"N:\Library\Design Software\Autodesk\Revit\Dynamo\packages\All Version"
source_base = r"N:\Library\Design Software\Autodesk\Revit\Dynamo\packages"
source_dir_year = os.path.join(source_base, "Revit {}".format(exe_year))
dest_base = os.path.expandvars("%appdata%")
dest_dir = _find_latest_dynamo_packages_dir(dest_base)

if exe_year < 2024:
    print("This tool is intended for Revit 2024+ only.")
    raise SystemExit
if not os.path.isdir(source_dir_year):
    print("Source packages folder not found: {}".format(source_dir_year))
    raise SystemExit
if not dest_dir:
    print("Could not find Dynamo packages folder under AppData.")
    raise SystemExit

if overwrite.lower() == "y" and os.path.exists(dest_dir):
	delete_files_in_folder(dest_dir)

if overwrite.lower() == "y" or not os.path.exists(dest_dir):
	shutil.copytree(source_dir_year, dest_dir, dirs_exist_ok=True)
	shutil.copytree(source_dir, dest_dir, dirs_exist_ok=True)
else:
	for root, dirs, files in os.walk(source_dir_year):
		relative_path = os.path.relpath(root, source_dir_year)
		destination = os.path.join(dest_dir, relative_path)
		os.makedirs(destination, exist_ok=True)
		for file in files:
			src_file = os.path.join(root, file)
			dst_file = os.path.join(destination, file)
			if not os.path.exists(dst_file):
				shutil.copy2(src_file, dst_file)
	# Always include shared packages
	shutil.copytree(source_dir, dest_dir, dirs_exist_ok=True)
