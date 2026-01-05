# import pyrevit libraries
import clr
from pyrevit import forms,script
from WWP_msgUtils import *

# check if notifications are disabled
if msgUtils_muted():
	script.exit()

# Get icon file (doesn't work)
# curPath = script.get_script_path()
# remPath = curPath.split('WWPTools.tab')[0]
# icoFile = remPath + r'bin\Graphics\ico256_WWP.ico'

# Display the message to the user
forms.toast("Toolbar has been loaded!","WWP_Tool",appid="WWP Architects + Planners",actions=None)
