import os
import shutil

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

source_dir = "N:\\Library\\Design Software\\Autodesk\\Revit\\Dynamo\\packages\\All Version"
source_dir1_2020 = "N:\\Library\\Design Software\\Autodesk\\Revit\\Dynamo\\packages\\Revit 2020"
source_dir1_2022 = "N:\\Library\\Design Software\\Autodesk\\Revit\\Dynamo\\packages\\Revit 2022"
source_dir1_2023 = "N:\\Library\\Design Software\\Autodesk\\Revit\\Dynamo\\packages\\Revit 2023"
dest_dir_2020 = os.path.expandvars("%appdata%") + "\\Dynamo\\Dynamo Revit\\2.3\\packages"
dest_dir_2022 = os.path.expandvars("%appdata%") + "\\Dynamo\\Dynamo Revit\\2.12\\packages"
dest_dir_2023 = os.path.expandvars("%appdata%") + "\\Dynamo\\Dynamo Revit\\2.16\\packages"

if overwrite.lower() == "y":
    delete_files_in_folder(dest_dir_2020)
    delete_files_in_folder(dest_dir_2022)
    delete_files_in_folder(dest_dir_2023)
    
if overwrite.lower() == "y" or not os.path.exists(dest_dir_2020):
    shutil.copytree(source_dir1_2020, dest_dir_2020, dirs_exist_ok=True)
    shutil.copytree(source_dir, dest_dir_2020, dirs_exist_ok=True)
else:
    for root, dirs, files in os.walk(source_dir1_2020):
        relative_path = os.path.relpath(root, source_dir1_2020)
        destination = os.path.join(dest_dir_2020, relative_path)
        os.makedirs(destination, exist_ok=True)
        for file in files:
            src_file = os.path.join(root, file)
            dst_file = os.path.join(destination, file)
            if not os.path.exists(dst_file):
                shutil.copy2(src_file, dst_file)

if overwrite.lower() == "y" or not os.path.exists(dest_dir_2022):
    shutil.copytree(source_dir1_2022, dest_dir_2022, dirs_exist_ok=True)
    shutil.copytree(source_dir, dest_dir_2022, dirs_exist_ok=True)
else:
    for root, dirs, files in os.walk(source_dir1_2022):
        relative_path = os.path.relpath(root, source_dir1_2022)
        destination = os.path.join(dest_dir_2022, relative_path)
        os.makedirs(destination, exist_ok=True)
        for file in files:
            src_file = os.path.join(root, file)
            dst_file = os.path.join(destination, file)
            if not os.path.exists(dst_file):
                shutil.copy2(src_file, dst_file)



if overwrite.lower() == "y" or not os.path.exists(dest_dir_2023):
    shutil.copytree(source_dir1_2023, dest_dir_2023, dirs_exist_ok=True)
    shutil.copytree(source_dir, dest_dir_2023, dirs_exist_ok=True)
else:
    for root, dirs, files in os.walk(source_dir1_2023):
        relative_path = os.path.relpath(root, source_dir1_2023)
        destination = os.path.join(dest_dir_2023, relative_path)
        os.makedirs(destination, exist_ok=True)
        for file in files:
            src_file = os.path.join(root, file)
            dst_file = os.path.join(destination, file)
            if not os.path.exists(dst_file):
                shutil.copy2(src_file, dst_file)
