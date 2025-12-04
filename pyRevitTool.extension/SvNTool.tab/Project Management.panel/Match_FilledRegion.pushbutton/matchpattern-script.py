import clr
clr.AddReference('RevitAPI')
import System
from System import Guid, Array
from System.Collections.Generic import *
from System.IO import Path
from Autodesk.Revit import DB
from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document
class FamilyOption(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source = FamilySource.Family
        overwriteParameterValues = True
        return True

# Find all Filled Region Type, and create a dictionary with it
modelFilledRegionTypes = dict()
filledinFamily = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).ToElements()
for e in filledinFamily:
    modelFilledRegionTypes[e.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()] = e
print(str(modelFilledRegionTypes))

# Find all loaded families
elements = DB.FilteredElementCollector(doc).OfClass(DB.Family).ToElements()
elementName=[]
for en in elements:
    eName= en.Name
    elementName.append(eName)
print(elementName)

# Initialize report list
report = []

# Loop on all loaded families

famdocs = []

for f in elements:
    if f.IsEditable:
        fName=f.Name
        famdocs.append(doc.EditFamily(f))
        print(fName)

for family in famdocs:
    #print("working on family:"+ family.Name)
    familyDoc = family
    familyEdited = False
    print("--Editing family...." + familyDoc.Title)
    # Find all Filled Region Type in the family
    filledinFamily = DB.FilteredElementCollector(familyDoc).OfClass(DB.FilledRegionType).ToElements()
    print(str(filledinFamily))
    with Transaction(familyDoc, "Edit Filled Region Type") as famTx:
        famTx.Start()
        for filledRegionType in filledinFamily:
            if filledRegionType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() in modelFilledRegionTypes:
                refFilledRegion = modelFilledRegionTypes[filledRegionType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()]
                print(refFilledRegion)
                # Change the foreground pattern color
                filledRegionType.ForegroundPatternColor = refFilledRegion.ForegroundPatternColor
                familyEdited = True
                print("foreground pattern color_checked")

                # Change the Background
                filledRegionType.BackgroundPatternColor = refFilledRegion.BackgroundPatternColor
                familyEdited = True
                print("background pattern color_checked" )

                # Change line weight
                filledRegionType.LineWeight = refFilledRegion.LineWeight
                familyEdited = True
                print("lineweight_checked")

                # Change fill pattern
                filledRegionType.BackgroundPatternId = refFilledRegion.BackgroundPatternId
                familyEdited = True
                print("filed pattern background_checked")                            
                # Change fill pattern
                try:
                    filledRegionType.ForegroundPatternId = refFilledRegion.ForegroundPatternId
                except:
                    newPat = FillPatternElement.Create(familyDoc, doc.GetElement(refFilledRegion.ForegroundPatternId).GetFillPattern())
                    filledRegionType.ForegroundPatternId = newPat.Id
                familyEdited = True
                print("fill pattern foreground_check " + str(filledRegionType.ForegroundPatternId))
                print("Current family"+ familyDoc.Title + " is done! ")
        famTx.Commit()

        if familyEdited:
            familyDoc.LoadFamily(doc, FamilyOption())
            familyDoc.Close(False)
        else:
            familyDoc.Close(False)

# Print report
for entry in report:
	print(entry)