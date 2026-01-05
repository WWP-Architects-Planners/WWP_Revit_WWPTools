# import pyrevit libraries
import clr
from pyrevit import forms,script
from SvN_msgUtils import *

# check if notifications are disabled
if msgUtils_muted():
	script.exit()

# Get icon file (doesn't work)
# curPath = script.get_script_path()
# remPath = curPath.split('SvNTool.tab')[0]
# icoFile = remPath + r'bin\Graphics\ico256_SvN.ico'

# Display the message to the user
forms.toast("Toolbar has been loaded!","SvN_Tool",appid="SvN Architects + Planners",actions=None)