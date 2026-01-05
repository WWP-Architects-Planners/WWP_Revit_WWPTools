# mute check
from pyrevit import script
from WWP_msgUtils import *

# import libraries
from datetime import datetime

# get the time
time = datetime.now()
dt_f = time.strftime("%Y-%m-%d %H-%M-%S")

# try to write data
try:
	script.store_data("Sync start", dt_f, this_project=True)
except:
	pass