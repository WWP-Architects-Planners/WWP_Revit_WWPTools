import clr
import os
from Autodesk.Revit.DB import ModelPathUtils, OpenOptions, DetachFromCentralOption, SaveAsOptions

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')

import RevitServices
from RevitServices.Persistence import DocumentManager
from pyrevit import forms


app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document

# Display folder dialog to select folder


def openDoc(filePath):
    if filePath is not None:
        modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath)
        options = OpenOptions()
        options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
        docOnDisk = app.OpenDocumentFile(modelPath, options)
        # Get current Revit version
        version = app.VersionNumber
        # Append RevitVersion to the end of the filename
        savePath = filePath[:-4] + "_R" + str(version)[-2::] + ".rvt"
        saveOptions = SaveAsOptions()
        saveOptions.OverwriteExistingFile = True
        docOnDisk.SaveAs(savePath, saveOptions)
        docOnDisk.Close(True)
        
def openFamily(filePath):
    if filePath is not None:
        modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath)
        options = OpenOptions()
        #options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
        docOnDisk = app.OpenDocumentFile(modelPath, options)
        # Get current Revit version
        version = app.VersionNumber
        # Append RevitVersion to the end of the filename
        savePath = filePath[:-4] + "_R" + str(version)[-2::] + ".rfa"
        saveOptions = SaveAsOptions()
        saveOptions.OverwriteExistingFile = True
        docOnDisk.SaveAs(savePath, saveOptions)
        docOnDisk.Close(True)

# Display folder dialog to select folder
folderpath = forms.pick_folder()        

# Loop through each file in the folder
for file in os.listdir(folderpath):
    if file.endswith('.rvt'):
        filepath = os.path.join(folderpath, file)
        openDoc(filepath)
    elif file.endswith('.rfa'):
        filepath = os.path.join(folderpath, file)
        openFamily(filepath)