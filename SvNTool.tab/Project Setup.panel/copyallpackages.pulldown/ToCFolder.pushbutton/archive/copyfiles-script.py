import os
import shutil
import stat

source_dir = r"N:\Library\Design Software\Autodesk\Revit\Dynamo\packages"
destination_dir = r"C:\dynpackages"

# Check if the destination directory exists, and if not, create it
if not os.path.exists(destination_dir):
    os.makedirs(destination_dir)

for root, dirs, files in os.walk(source_dir):
    # Calculate the corresponding destination directory
    relative_path = os.path.relpath(root, source_dir)
    dest_dir = os.path.join(destination_dir, relative_path)
    print("Copying to {}".format(dest_dir))
    # Create the corresponding destination directory
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)

    for file in files:
        source_file = os.path.join(root, file)
        dest_file = os.path.join(dest_dir, file)

        # Check if the file exists in the destination directory
        if os.path.exists(dest_file):
            # Check if the source file is newer (based on modified date)
            if os.path.getmtime(source_file) > os.path.getmtime(dest_file):
                #print("Copying {} to {}".format(source_file, dest_file))
                shutil.copy2(source_file, dest_file)
            else:
                print("Skipping {} (no changes)".format(source_file))
        else:
            print("Copying {} to {}".format(source_file, dest_file))
            shutil.copy2(source_file, dest_file)

        # Ensure copied files are not read-only
        if os.path.exists(dest_file):
            os.chmod(dest_file, stat.S_IWRITE)

print('''







        Installation complete!!!!!  ''')
