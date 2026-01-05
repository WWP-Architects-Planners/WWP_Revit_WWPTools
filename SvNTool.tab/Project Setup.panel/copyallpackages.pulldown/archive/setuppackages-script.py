import subprocess

def copy_packages():
    batch_file = r"N:\Library\Design Software\Autodesk\Revit\Dynamo\packages\copypackages.bat"
    subprocess.call([batch_file])

copy_packages()