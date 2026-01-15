# import libraries
import clr
import os

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Drawing import Bitmap, Icon, Point, Size
from System.Windows.Forms import (
	AnchorStyles,
	Button,
	CheckedListBox,
	DialogResult,
	Form,
	FormStartPosition,
	Label,
	ListBox,
	PictureBox,
	PictureBoxSizeMode,
	SelectionMode,
)

_LOGO_ICON = None
_LOGO_BITMAP = None

def _logo_path():
	cur_dir = os.path.dirname(os.path.abspath(__file__))
	return os.path.abspath(os.path.join(cur_dir, "WWPtools-logo.png"))

def _load_logo_icon():
	global _LOGO_ICON
	global _LOGO_BITMAP
	if _LOGO_ICON is not None:
		return _LOGO_ICON
	logo_path = _logo_path()
	if not os.path.isfile(logo_path):
		return None
	try:
		_LOGO_BITMAP = Bitmap(logo_path)
		_LOGO_ICON = Icon.FromHandle(_LOGO_BITMAP.GetHicon())
		return _LOGO_ICON
	except Exception:
		return None

def _load_logo_bitmap():
	global _LOGO_BITMAP
	if _LOGO_BITMAP is not None:
		return _LOGO_BITMAP
	logo_path = _logo_path()
	if not os.path.isfile(logo_path):
		return None
	try:
		_LOGO_BITMAP = Bitmap(logo_path)
		return _LOGO_BITMAP
	except Exception:
		return None

def _apply_logo(form, show_inline=True):
	icon = _load_logo_icon()
	if icon:
		form.Icon = icon
	if show_inline:
		bitmap = _load_logo_bitmap()
		if bitmap:
			pic = PictureBox()
			pic.Image = bitmap
			pic.SizeMode = PictureBoxSizeMode.Zoom
			pic.Size = Size(40, 40)
			pic.Location = Point(form.Width - 74, 8)
			pic.Anchor = AnchorStyles.Top | AnchorStyles.Right
			form.Controls.Add(pic)

def _message_form(message, title, show_cancel=False, width=520, height=220):
	form = Form()
	form.Text = title
	form.StartPosition = FormStartPosition.CenterScreen
	form.Size = Size(width, height)
	form.MinimizeBox = False
	form.MaximizeBox = False
	_apply_logo(form)

	label = Label()
	label.Text = message
	label.Location = Point(12, 12)
	label.Size = Size(width - 40, height - 90)
	label.AutoSize = False
	label.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
	form.Controls.Add(label)

	ok_btn = Button()
	ok_btn.Text = "OK"
	ok_btn.DialogResult = DialogResult.OK
	ok_btn.Location = Point(width - 204, height - 72)
	form.Controls.Add(ok_btn)

	if show_cancel:
		cancel_btn = Button()
		cancel_btn.Text = "Cancel"
		cancel_btn.DialogResult = DialogResult.Cancel
		cancel_btn.Location = Point(width - 112, height - 72)
		form.Controls.Add(cancel_btn)
		form.CancelButton = cancel_btn

	form.AcceptButton = ok_btn
	return form

# show a simple alert dialog
def uiUtils_alert(message, title="Message"):
	form = _message_form(message, title, show_cancel=False)
	form.ShowDialog()

# show a confirm dialog
def uiUtils_confirm(message, title="Confirm"):
	form = _message_form(message, title, show_cancel=True)
	return form.ShowDialog() == DialogResult.OK

# show a list picker and return selected indices
def uiUtils_select_indices(items, title="Select Items", prompt="Select items:", multiselect=True, width=980, height=540):
	form = Form()
	form.Text = title
	form.StartPosition = FormStartPosition.CenterScreen
	form.Size = Size(width, height)
	form.MinimizeBox = False
	form.MaximizeBox = False
	_apply_logo(form, show_inline=True)

	info = Label()
	info.Text = prompt
	info.Location = Point(12, 12)
	info.AutoSize = True
	form.Controls.Add(info)

	listbox = None
	checklist = None
	if multiselect:
		checklist = CheckedListBox()
		checklist.Location = Point(12, 36)
		checklist.Size = Size(width - 40, height - 120)
		checklist.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
		checklist.CheckOnClick = True
		for item in items:
			checklist.Items.Add(item)
		form.Controls.Add(checklist)
	else:
		listbox = ListBox()
		listbox.Location = Point(12, 36)
		listbox.Size = Size(width - 40, height - 120)
		listbox.SelectionMode = SelectionMode.One
		listbox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
		for item in items:
			listbox.Items.Add(item)
		form.Controls.Add(listbox)

	ok_btn = Button()
	ok_btn.Text = "OK"
	ok_btn.DialogResult = DialogResult.OK
	ok_btn.Location = Point(width - 204, height - 72)
	form.Controls.Add(ok_btn)

	cancel_btn = Button()
	cancel_btn.Text = "Cancel"
	cancel_btn.DialogResult = DialogResult.Cancel
	cancel_btn.Location = Point(width - 112, height - 72)
	form.Controls.Add(cancel_btn)

	form.AcceptButton = ok_btn
	form.CancelButton = cancel_btn

	if form.ShowDialog() != DialogResult.OK:
		return []

	selected = []
	if multiselect and checklist:
		for index in checklist.CheckedIndices:
			selected.append(index)
	elif listbox:
		for index in listbox.SelectedIndices:
			selected.append(index)
	return selected
