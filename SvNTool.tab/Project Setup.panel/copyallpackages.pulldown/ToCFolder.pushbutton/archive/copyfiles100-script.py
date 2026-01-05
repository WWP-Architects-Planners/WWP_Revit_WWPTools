#! python3
import os
import shutil
import logging

# Set up logging to a file
log_file_path = r"C:\dynpackages\log.txt"
logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Step 1: Create "C:\dynpackages" if it doesn't exist
dynpackages_path = r"C:\dynpackages"
if not os.path.exists(dynpackages_path):
    os.makedirs(dynpackages_path)
    logging.info(f"Created directory: {dynpackages_path}")

# Step 2: Try deleting "C:\dynpackages" if it exists
try:
    shutil.rmtree(dynpackages_path)
    logging.info(f"Deleted directory: {dynpackages_path}")
except Exception as e:
    logging.warning(f"Unable to delete directory: {e}")

# Step 3: Copy folders from "N:\Library\Design Software\Autodesk\Revit\Dynamo\packages" to "C:\dynpackages"
source_path = r"N:\Library\Design Software\Autodesk\Revit\Dynamo\packages"

if os.path.exists(source_path):
    for item in os.listdir(source_path):
        source_item = os.path.join(source_path, item)
        destination_item = os.path.join(dynpackages_path, item)

        try:
            if os.path.isdir(source_item):
                shutil.copytree(source_item, destination_item, symlinks=True, ignore_dangling_symlinks=True)
                logging.info(f"Copied directory: {source_item} to {destination_item}")
            else:
                shutil.copy2(source_item, destination_item)
                logging.info(f"Copied file: {source_item} to {destination_item}")
        except Exception as e:
            logging.warning(f"Unable to copy {source_item} to {destination_item}: {e}")
else:
    logging.warning(f"Source directory does not exist: {source_path}")
