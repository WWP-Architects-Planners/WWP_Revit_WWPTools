# Prepare for converters
from pyrevit import revit,DB,script
import WWP_uiUtils as ui

# function for getting element from string
def delUtils_getFromKey(s):
	splitName = s.rsplit("[")[-1]
	strId = splitName.replace("]","")
	eleId = DB.ElementId(int(strId))
	return eleId

# function for pb length
def delUtils_pbStep(lst, divs = 10):
	try:
		stepRound = int(len(lst)/divs)
		if stepRound < 1:
			return 1
		else:
			return stepRound
	except:
		return 1

def delUtils_pbCancelled(pb,pbMsg="Script cancelled.",pbTitle="Script cancelled."):
	if pb.cancelled:
		ui.uiUtils_alert(pbMsg, title= pbTitle)
		script.exit()

# function for deletion
def delUtils_delEle(e,myDoc=revit.doc):
	try:
		myDoc.Delete(e.Id)
		return 1
	except:
		try:
			myDoc.Delete(e)
			return 1
		except:
			return 0

# delete elements with reporting
def delUtils_delEles(eles,myDoc=revit.doc,pbTitle="Deleting elements..."):
	pbTotal = len(eles)
	del_pass = 0
	with revit.Transaction('Delete elements'):
		for e in eles:
			del_pass += delUtils_delEle(e)
	form_message = str(del_pass) + "/" + str(pbTotal) + " elements successfully deleted."
	ui.uiUtils_alert(form_message, title="Deletion completed")
