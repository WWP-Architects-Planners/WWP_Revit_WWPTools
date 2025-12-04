import os
import shutil
import ctypes

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

    source_dirs = ["Revit 2020", "Revit 2022", "Revit 2023", "Revit 2024"]
    dest_dirs = ["2.3\\packages", "2.12\\packages", "2.16\\packages", "2.18\\packages"]

    for i in range(len(source_dirs)):
        source_dir = os.path.join(source_base, source_dirs[i])
        dest_dir = os.path.join(dest_base, r"Dynamo\Dynamo Revit", dest_dirs[i])

        print("Copying files from: \n{}\n".format(source_dir))
        print("To destination: \n{}\n".format(dest_dir))

        if os.path.exists(dest_dir):
            print("Deleting existing files in: \n{}\n".format(dest_dir))
            shutil.rmtree(dest_dir, ignore_errors=True)
        copy_files(source_dir, dest_dir)
        print("Files copied successfully.\n")

        print("Copying extra folder from: \n{}\n".format(source_all_version))
        print("To destination: \n{}\n".format(dest_dir))
        copy_files(source_all_version, dest_dir)
        print("Folder copied\nsuccessfully.\n")
else:
    print("Copy process cancelled.")
print("\nAll Process is done")
