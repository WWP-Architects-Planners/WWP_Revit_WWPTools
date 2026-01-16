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
	CheckBox,
	DialogResult,
	Form,
	FormStartPosition,
	Label,
	ListBox,
	PictureBox,
	PictureBoxSizeMode,
	RadioButton,
	TextBox,
	ComboBox,
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


def uiUtils_select_items_with_mode(
	items,
	title="Select Items",
	prompt="Select items:",
	mode_labels=("Option A", "Option B"),
	default_mode=0,
	prechecked_indices=None,
	width=720,
	height=620,
):
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

	checklist = CheckedListBox()
	checklist.Location = Point(12, 36)
	checklist.Size = Size(width - 40, height - 180)
	checklist.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
	checklist.CheckOnClick = True
	for item in items:
		checklist.Items.Add(item)
	if prechecked_indices:
		for idx in prechecked_indices:
			if 0 <= idx < checklist.Items.Count:
				checklist.SetItemChecked(idx, True)
	form.Controls.Add(checklist)

	rb_left = RadioButton()
	rb_left.Text = mode_labels[0] if mode_labels else "Option A"
	rb_left.Location = Point(12, height - 132)
	rb_left.Checked = default_mode == 0
	form.Controls.Add(rb_left)

	rb_right = RadioButton()
	rb_right.Text = mode_labels[1] if len(mode_labels or []) > 1 else "Option B"
	rb_right.Location = Point(160, height - 132)
	rb_right.Checked = default_mode == 1
	form.Controls.Add(rb_right)

	ok_btn = Button()
	ok_btn.Text = "Export"
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
		return [], None

	selected = []
	for index in checklist.CheckedIndices:
		selected.append(index)
	mode = 0 if rb_left.Checked else 1
	return selected, mode


def _safe_iter(items):
	if items is None:
		return []
	try:
		return list(items)
	except TypeError:
		return []


def uiUtils_select_sheet_renumber_inputs(
	categories=None,
	print_sets=None,
	title="Renumber Sheets",
	category_label="Choose Sheet Category",
	printset_label="Choose Print Set",
	starting_label="Starting Number",
	cancel_text="Cancel",
	width=520,
	height=320,
):
	form = Form()
	form.Text = title
	form.StartPosition = FormStartPosition.CenterScreen
	form.Size = Size(width, height)
	form.MinimizeBox = False
	form.MaximizeBox = False
	_apply_logo(form, show_inline=True)

	current_y = 20

	category_items = _safe_iter(categories)
	printset_items = _safe_iter(print_sets)

	category_combo = None
	if categories is not None:
		label_category = Label()
		label_category.Text = category_label
		label_category.Location = Point(12, current_y)
		label_category.AutoSize = True
		form.Controls.Add(label_category)

		category_combo = ComboBox()
		category_combo.Location = Point(12, current_y + 24)
		category_combo.Size = Size(width - 40, 24)
		for item in category_items:
			category_combo.Items.Add(item)
		if category_combo.Items.Count > 0:
			category_combo.SelectedIndex = 0
		form.Controls.Add(category_combo)
		current_y += 64

	printset_combo = None
	if print_sets is not None:
		label_printset = Label()
		label_printset.Text = printset_label
		label_printset.Location = Point(12, current_y)
		label_printset.AutoSize = True
		form.Controls.Add(label_printset)

		printset_combo = ComboBox()
		printset_combo.Location = Point(12, current_y + 24)
		printset_combo.Size = Size(width - 40, 24)
		for item in printset_items:
			printset_combo.Items.Add(item)
		if printset_combo.Items.Count > 0:
			printset_combo.SelectedIndex = 0
		form.Controls.Add(printset_combo)
		current_y += 64

	label_start = Label()
	label_start.Text = starting_label
	label_start.Location = Point(12, current_y)
	label_start.AutoSize = True
	form.Controls.Add(label_start)

	start_input = TextBox()
	start_input.Location = Point(12, current_y + 24)
	start_input.Size = Size(width - 40, 24)
	form.Controls.Add(start_input)

	ok_btn = Button()
	ok_btn.Text = "Set Values"
	ok_btn.DialogResult = DialogResult.OK
	ok_btn.Location = Point(width - 204, height - 72)
	form.Controls.Add(ok_btn)

	cancel_btn = Button()
	cancel_btn.Text = cancel_text or "Cancel"
	cancel_btn.DialogResult = DialogResult.Cancel
	cancel_btn.Location = Point(width - 112, height - 72)
	form.Controls.Add(cancel_btn)

	form.AcceptButton = ok_btn
	form.CancelButton = cancel_btn

	if form.ShowDialog() != DialogResult.OK:
		return None

	category_value = ""
	if category_combo is not None:
		category_value = category_combo.SelectedItem if category_combo.SelectedItem else ""
	printset_value = ""
	if printset_combo is not None:
		printset_value = printset_combo.SelectedItem if printset_combo.SelectedItem else ""
	starting_value = start_input.Text or ""
	return {
		"category": category_value,
		"printset": printset_value,
		"starting_number": starting_value,
	}


def uiUtils_select_sheet_renumber_inputs_with_list(
	items,
	title="Renumber Sheets",
	prompt="Select sheets to renumber:",
	starting_label="Starting Number",
	cancel_text="Cancel",
	width=980,
	height=620,
):
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

	checklist = CheckedListBox()
	checklist.Location = Point(12, 36)
	checklist.Size = Size(width - 40, height - 190)
	checklist.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
	checklist.CheckOnClick = True
	for item in _safe_iter(items):
		checklist.Items.Add(item)
	form.Controls.Add(checklist)

	label_start = Label()
	label_start.Text = starting_label
	label_start.Location = Point(12, height - 136)
	label_start.AutoSize = True
	form.Controls.Add(label_start)

	start_input = TextBox()
	start_input.Location = Point(12, height - 112)
	start_input.Size = Size(width - 40, 24)
	form.Controls.Add(start_input)

	ok_btn = Button()
	ok_btn.Text = "Set Values"
	ok_btn.DialogResult = DialogResult.OK
	ok_btn.Location = Point(width - 204, height - 72)
	form.Controls.Add(ok_btn)

	cancel_btn = Button()
	cancel_btn.Text = cancel_text or "Cancel"
	cancel_btn.DialogResult = DialogResult.Cancel
	cancel_btn.Location = Point(width - 112, height - 72)
	form.Controls.Add(cancel_btn)

	form.AcceptButton = ok_btn
	form.CancelButton = cancel_btn

	if form.ShowDialog() != DialogResult.OK:
		return None

	selected = []
	for index in checklist.CheckedIndices:
		selected.append(index)

	return {
		"selected_indices": selected,
		"starting_number": start_input.Text or "",
	}


def uiUtils_viewname_replace_inputs(
	title="Replace View Name",
	find_label="Find",
	replace_label="Replace",
	prefix_label="Prefix",
	suffix_label="Suffix",
	cancel_text="Cancel",
	width=520,
	height=320,
):
	form = Form()
	form.Text = title
	form.StartPosition = FormStartPosition.CenterScreen
	form.Size = Size(width, height)
	form.MinimizeBox = False
	form.MaximizeBox = False
	_apply_logo(form, show_inline=True)

	current_y = 20

	label_find = Label()
	label_find.Text = find_label
	label_find.Location = Point(12, current_y)
	label_find.AutoSize = True
	form.Controls.Add(label_find)

	find_input = TextBox()
	find_input.Location = Point(12, current_y + 24)
	find_input.Size = Size(width - 40, 24)
	form.Controls.Add(find_input)
	current_y += 56

	label_replace = Label()
	label_replace.Text = replace_label
	label_replace.Location = Point(12, current_y)
	label_replace.AutoSize = True
	form.Controls.Add(label_replace)

	replace_input = TextBox()
	replace_input.Location = Point(12, current_y + 24)
	replace_input.Size = Size(width - 40, 24)
	form.Controls.Add(replace_input)
	current_y += 56

	label_prefix = Label()
	label_prefix.Text = prefix_label
	label_prefix.Location = Point(12, current_y)
	label_prefix.AutoSize = True
	form.Controls.Add(label_prefix)

	prefix_input = TextBox()
	prefix_input.Location = Point(12, current_y + 24)
	prefix_input.Size = Size(width - 40, 24)
	form.Controls.Add(prefix_input)
	current_y += 56

	label_suffix = Label()
	label_suffix.Text = suffix_label
	label_suffix.Location = Point(12, current_y)
	label_suffix.AutoSize = True
	form.Controls.Add(label_suffix)

	suffix_input = TextBox()
	suffix_input.Location = Point(12, current_y + 24)
	suffix_input.Size = Size(width - 40, 24)
	form.Controls.Add(suffix_input)

	ok_btn = Button()
	ok_btn.Text = "Apply"
	ok_btn.DialogResult = DialogResult.OK
	ok_btn.Location = Point(width - 204, height - 72)
	form.Controls.Add(ok_btn)

	cancel_btn = Button()
	cancel_btn.Text = cancel_text or "Cancel"
	cancel_btn.DialogResult = DialogResult.Cancel
	cancel_btn.Location = Point(width - 112, height - 72)
	form.Controls.Add(cancel_btn)

	form.AcceptButton = ok_btn
	form.CancelButton = cancel_btn

	if form.ShowDialog() != DialogResult.OK:
		return None

	return {
		"find": find_input.Text or "",
		"replace": replace_input.Text or "",
		"prefix": prefix_input.Text or "",
		"suffix": suffix_input.Text or "",
	}


def uiUtils_duplicate_sheet_inputs(
	items,
	title="Sheet Duplicator",
	prompt="Select sheets to duplicate:",
	options_label="Duplicate Options",
	duplicate_with_views_label="Duplicate with Views",
	prefix_label="Sheet Number Prefix",
	suffix_label="Sheet Number Suffix",
	cancel_text="Cancel",
	width=980,
	height=700,
):
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

	checklist = CheckedListBox()
	checklist.Location = Point(12, 36)
	checklist.Size = Size(width - 40, height - 300)
	checklist.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
	checklist.CheckOnClick = True
	for item in _safe_iter(items):
		checklist.Items.Add(item)
	form.Controls.Add(checklist)

	views_checkbox = CheckBox()
	views_checkbox.Text = duplicate_with_views_label
	views_checkbox.Location = Point(12, height - 250)
	views_checkbox.Checked = True
	form.Controls.Add(views_checkbox)

	label_options = Label()
	label_options.Text = options_label
	label_options.Location = Point(12, height - 220)
	label_options.AutoSize = True
	form.Controls.Add(label_options)

	options_combo = ComboBox()
	options_combo.Location = Point(12, height - 196)
	options_combo.Size = Size(width - 40, 24)
	for item in ("Duplicate Views", "Duplicate Views w/Details", "Duplicate Views AsDependent"):
		options_combo.Items.Add(item)
	if options_combo.Items.Count > 0:
		options_combo.SelectedIndex = 0
	form.Controls.Add(options_combo)

	label_prefix = Label()
	label_prefix.Text = prefix_label
	label_prefix.Location = Point(12, height - 164)
	label_prefix.AutoSize = True
	form.Controls.Add(label_prefix)

	prefix_input = TextBox()
	prefix_input.Location = Point(12, height - 140)
	prefix_input.Size = Size(width - 40, 24)
	form.Controls.Add(prefix_input)

	label_suffix = Label()
	label_suffix.Text = suffix_label
	label_suffix.Location = Point(12, height - 108)
	label_suffix.AutoSize = True
	form.Controls.Add(label_suffix)

	suffix_input = TextBox()
	suffix_input.Location = Point(12, height - 84)
	suffix_input.Size = Size(width - 40, 24)
	form.Controls.Add(suffix_input)

	ok_btn = Button()
	ok_btn.Text = "Duplicate"
	ok_btn.DialogResult = DialogResult.OK
	ok_btn.Location = Point(width - 204, height - 52)
	form.Controls.Add(ok_btn)

	cancel_btn = Button()
	cancel_btn.Text = cancel_text or "Cancel"
	cancel_btn.DialogResult = DialogResult.Cancel
	cancel_btn.Location = Point(width - 112, height - 52)
	form.Controls.Add(cancel_btn)

	form.AcceptButton = ok_btn
	form.CancelButton = cancel_btn

	if form.ShowDialog() != DialogResult.OK:
		return None

	selected = []
	for index in checklist.CheckedIndices:
		selected.append(index)

	return {
		"selected_indices": selected,
		"duplicate_with_views": views_checkbox.Checked,
		"duplicate_option": options_combo.SelectedIndex if options_combo.SelectedIndex >= 0 else 0,
		"prefix": prefix_input.Text or "",
		"suffix": suffix_input.Text or "",
	}
