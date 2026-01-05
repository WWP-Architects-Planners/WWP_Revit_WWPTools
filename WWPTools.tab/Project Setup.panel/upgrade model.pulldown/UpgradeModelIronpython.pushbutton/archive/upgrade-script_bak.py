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

def detach_and_save(doc, save_path):
    options = SaveAsOptions()
    options.OverwriteExistingFile = False

    # Detach and preserve worksets
    detach_options = DetachFromCenoqguuefqxntralOption.DetachAndPreserveWorksets
    detached_doc = doc.UnloadWorksharing(detach_options)
    
    if success:
        # Save the detached document
        detached_doc.SaveAs(save_path, options)
        detached_doc.Close(True)
    else:
        raise Exception("Failed to detach and save the document.")

def open_doc(filePath):
    if filePath is not None:
        model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath)
        options = OpenOptions()
        doc_on_disk = app.OpenDocumentFile(model_path, options)
        
        # Get current Revit version
        version = app.VersionNumber
        
        # Append RevitVersion to the end of the filename
        save_path = filePath[:-4] + "_R" + str(version)[-2::] + ".rvt"
        
        # Detach and save the document
        detach_and_save(doc_on_disk, save_path)

# Display folder dialog to select folder
folder_path = forms.pick_folder()

# Loop through each file in the folder
for file in os.listdir(folder_path):
    if file.endswith('.rvt'):
        file_path = os.path.join(folder_path, file)
        open_doc(file_path)
    elif file.endswith('.rfa'):
        file_path = os.path.join(folder_path, file)
        open_doc(file_path)
