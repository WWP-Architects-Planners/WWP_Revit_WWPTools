# -*- coding: utf-8 -*-
# import pyrevit libraries
import os
import clr
from pyrevit import forms, script
from WWP_msgUtils import *
import json

try:
	import WWP_telemetry
	WWP_telemetry.track_app_init()
except Exception:
	pass

# ---------------------------------------------------------------------------
# Enable pyRevit Script Telemetry and process completed session files.
# pyRevit's C# ScriptExecutor fires telemetry after every script run -
# this is the only mechanism that catches IExternalCommand pyRevit buttons.
# On each startup we read the previous session's JSON file, filter for
# WWPTools entries, append to script_log.jsonl, and POST to Neon.
# ---------------------------------------------------------------------------
try:
	import glob
	import shutil
	import sys
	import socket
	from pyrevit import telemetry as _pyrvit_tel
	from pyrevit.userconfig import user_config as _pyr_cfg
	from pyrevit.loader import sessioninfo as _pyr_si

	_appdata    = os.environ.get("APPDATA", "")
	_wwp_dir    = os.path.join(_appdata, "pyRevit", "WWPTools")
	_log_path   = os.path.join(_wwp_dir, "script_log.jsonl")
	_done_dir   = os.path.join(_wwp_dir, "_processed_telemetry")
	if not os.path.isdir(_wwp_dir):
		os.makedirs(_wwp_dir)

	# Enable Script Telemetry in config (persists for all future sessions)
	if not _pyr_cfg.telemetry_status or not _pyr_cfg.telemetry_file_dir:
		_pyr_cfg.telemetry_status   = True
		_pyr_cfg.telemetry_file_dir = _wwp_dir
		_pyr_cfg.save_changes()

	# Also wire up the current session's telemetry file path immediately
	# so this session's runs are captured (not just future ones)
	try:
		from pyrevit import PYREVIT_FILE_PREFIX
		_sid = _pyr_si.get_session_uuid()
		_tel_filename = "{}_{}_telemetry.json".format(PYREVIT_FILE_PREFIX, _sid)
		_tel_filepath = os.path.join(_wwp_dir, _tel_filename)
		if not os.path.isfile(_tel_filepath):
			with open(_tel_filepath, "w") as _tf:
				_tf.write("[]")
		_pyrvit_tel.set_telemetry_state(True)
		_pyrvit_tel.set_telemetry_file_dir(_wwp_dir)
		_pyrvit_tel.set_telemetry_file_path(_tel_filepath)
	except BaseException:
		pass

	# Process completed session telemetry files (skip the current one)
	_current_tel = os.path.normpath(_pyrvit_tel.get_telemetry_file_path() or "")
	if not os.path.isdir(_done_dir):
		os.makedirs(_done_dir)

	_lib_path = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "lib"))
	if _lib_path not in sys.path:
		sys.path.insert(0, _lib_path)
	import WWP_telemetry as _wwp_tel

	_reported_signatures = set()
	for _tel_file in glob.glob(os.path.join(_wwp_dir, "*_telemetry.json")):
		if os.path.normpath(_tel_file) == _current_tel:
			continue
		try:
			with open(_tel_file, "r") as _fh:
				_data = json.load(_fh)
			_records = _data.get("value", []) if isinstance(_data, dict) else (_data or [])
			for _rec in _records:
				# Filter to WWPTools commands only
				if str(_rec.get("commandextension", "") or "").lower() != "wwp_revit_wwptools":
					continue
				_trace        = _rec.get("trace") or {}
				_success      = int(_rec.get("resultcode") or 0) == 0
				_err          = str(_trace.get("message", "") or "").strip() or None
				_script_name  = str(_rec.get("commandname", ""))
				_user_name    = str(_rec.get("username", ""))
				_machine_name = socket.gethostname()
				_revit_ver    = str(_rec.get("revit", ""))
				_doc_name     = str(_rec.get("docname", "") or "")
				_evt_type     = "error" if not _success else "tool_use"
				_details      = _wwp_tel._format_details_md(
					_script_name, _evt_type, _success, _user_name, _machine_name,
					_revit_ver, None, None, _doc_name, 0, _err,
				)
				_entry = {
					"logged_at":        str(_rec.get("exec_timestamp", "") or _rec.get("timestamp", "")),
					"user_name":        _user_name,
					"machine_name":     _machine_name,
					"script_name":      _script_name,
					"script_type":      "python",
					"revit_version":    _revit_ver,
					"success":          _success,
					"error_msg":        _err,
					"document_name":    _doc_name,
					"session_id":       str(_rec.get("sessionid", "") or ""),
					"event_type":       _evt_type,
					"pyrevit_version":  str(_rec.get("pyrevit", "") or ""),
					"wwptools_version": None,
					"details":          _details,
				}
				import io as _io
				with _io.open(_log_path, "a", encoding="utf-8") as _lf:
					_lf.write(json.dumps(_entry, ensure_ascii=False) + u"\n")
				if _wwp_tel.is_telemetry_enabled():
					_wwp_tel._worker(_entry)
					if _evt_type == "error" and _err:
						try:
							_dedup_key = (_script_name, _err)
							if _dedup_key not in _reported_signatures:
								_reported_signatures.add(_dedup_key)
								_wwp_tel._post_error_report(_entry)
						except Exception:
							pass
			shutil.move(_tel_file, os.path.join(_done_dir, os.path.basename(_tel_file)))
		except BaseException as _exc:
			try:
				import traceback as _tb
				_dbg = os.path.join(_wwp_dir, "telemetry_debug.txt")
				with open(_dbg, "a") as _df:
					_df.write("FILE: {}\nERROR: {}\n{}\n---\n".format(
						_tel_file, str(_exc), _tb.format_exc()))
			except Exception:
				pass
except BaseException:
	pass

# check if notifications are disabled
if msgUtils_muted():
	script.exit()

# Get icon file (doesn't work)
# curPath = script.get_script_path()
# remPath = curPath.split('WWPTools.tab')[0]
# icoFile = remPath + r'bin\Graphics\ico256_WWP.ico'

# Display the message to the user
icon_path = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "lib", "WWPtools-logo.png"))
toolbar_title = "WWP_Tool"
toolbar_msg = "Toolbar has been loaded!"

# ----------------------------------------------------
# Update check (GitHub latest release via WWP_update_service)
# ----------------------------------------------------
try:
	import WWP_update_service as _update_svc
	_pending_release = _update_svc.check_in_background()
	if _pending_release:
		toolbar_msg = "Toolbar loaded. Your version ({}) is outdated.".format(
			_update_svc.installed_version_text(),
		)
		_update_svc.show_update_prompt(_pending_release, offer_skip=True)
	else:
		_ver_text = _update_svc.installed_version_text()
		if _ver_text and _ver_text != "unknown":
			toolbar_msg = "Toolbar loaded. You are running the latest version ({})".format(_ver_text)
except Exception:
	pass

try:
	if os.path.exists(icon_path):
		forms.toaster.send_toast(
			toolbar_msg,
			title=toolbar_title,
			appid="WWP Architects + Planners",
			icon=icon_path,
			click=None,
			actions=None,
		)
	else:
		forms.toast(
			toolbar_msg,
			toolbar_title,
			appid="WWP Architects + Planners",
			actions=None,
		)
except Exception:
	pass
