# mute check
from pyrevit import script
from SvN_msgUtils import *

# check if notifications are disabled
if msgUtils_muted():
	script.exit()

# import libraries
from pyrevit import forms
from datetime import datetime
import math

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
		msg_title = "Completed, it synced faster than Lightning!!"
	elif e_mins > 5:
		msg_time = timeToString(e_mins) + " mins & " + timeToString(e_secs) + " secs"
		msg_title = "Synced but it took too long(check your model)"
	else:
		msg_time = timeToString(e_mins) + " mins & " + timeToString(e_secs) + " secs"
		msg_title = "Sync Completed"
	# Show toast message
	forms.toaster.send_toast(
        msg_time, 
        title = msg_title, 
        appid=None, 
        icon="N:\\Library\\Design Software\\Autodesk\\Revit\\Dynamo\\PYREVIT\\SVN LOGO.jpg", 
        click=None, 
        actions=None
       )
except:
	pass