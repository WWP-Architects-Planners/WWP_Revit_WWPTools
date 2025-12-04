import clr
import os
from Autodesk.Revit.DB import ModelPathUtils, OpenOptions, DetachFromCentralOption, SaveAsOptions

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')

import RevitServices
from RevitServices.Persistence import DocumentManager
from pyrevit import forms
from Autodesk.Revit.ApplicationServices import Application

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument