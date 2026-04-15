# mute check
from pyrevit import script
from WWP_msgUtils import *

# check if notifications are disabled
if msgUtils_muted():
	script.exit()

# import libraries
from pyrevit import forms
from datetime import datetime
import math
import os

# get the time
time = datetime.now()
dt_f2 = time.strftime("%Y-%m-%d %H-%M-%S")

# try to write data
# Get a datetime
dt_f1  = script.load_data("Sync start", this_project=True)

def getdhms(dt_f):
    date_str = dt_f.split(" ")[0]
    time_str = dt_f.split(" ")[1]
    day = date_str.split("-")[2]
    hour= time_str.split("-")[0]
    minute = time_str.split("-")[1]
    second = time_str.split("-")[2]
    timelst = [day,hour,minute,second]
    return timelst

def diff(la,lb):
    d=(int(lb[0])-int(la[0]))*60*60*24
    h=(int(lb[1])-int(la[1]))*60*60
    m=(int(lb[2])-int(la[2]))*60
    s=int(lb[3])-int(la[3])
    return d+h+m+s

def timeToString(s):
	unpad = str(s).split(".")[0]
	return unpad.rjust(2,"0")

# Break down to parts
before = getdhms(dt_f1)
after = getdhms(dt_f2)
msg = ""

try:
	elapsed = diff(before, after)
	e_mins  = math.floor(elapsed/60)
	e_secs  = elapsed % 60
	if e_mins < 1:
		msg_time = "Only " + timeToString(e_secs) + " secs"
		msg_title = "Sync complete"
	elif e_mins > 5:
		msg_time = timeToString(e_mins) + " mins & " + timeToString(e_secs) + " secs"
		msg_title = "Sync complete (slow)"
	else:
		msg_time = timeToString(e_mins) + " mins & " + timeToString(e_secs) + " secs"
		msg_title = "Sync complete"
	if e_mins >= 10:
		msg_time = msg_time + "\nModel health: https://svn-architects-planners-inc.gitbook.io/svn-guidebooks/w7kFyDX0kRTb27slSn93/wwp-technical-guidebook/section-2-or-revit/2.2-or-general-info/2.1.5-or-important-concepts/2.1.5.4-or-revit-health"
	# Show toast message
	icon_path = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "lib", "WWPtools-logo.png"))
	forms.toaster.send_toast(
        msg_time, 
        title = msg_title, 
        appid=None, 
        icon=icon_path, 
        click=None, 
        actions=None
       )
except:
	pass
