import clr
import os

clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from System.Drawing import Bitmap, Icon, Point, Size
from System.Windows.Forms import (
    AnchorStyles,
    Button,
    DialogResult,
    Form,
    FormStartPosition,
    GroupBox,
    Label,
    PictureBox,
    PictureBoxSizeMode,
    RadioButton,
    TextBox,
)

_LOGO_BITMAP = None
_LOGO_ICON = None


def _logo_path():
    cur_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(cur_dir, "WWPtools-logo.png"))


def _load_logo_icon():
    global _LOGO_BITMAP
    global _LOGO_ICON
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


class _UIForm(Form):
    def __init__(
        self,
        title,
        description,
        prefix_default,
        suffix_default,
        duplicate_options,
        default_index,
        button_text,
    ):
        Form.__init__(self)
        self.Text = title
        self.Width = 520
        self.Height = 320
        self.StartPosition = FormStartPosition.CenterScreen
        self.MinimizeBox = False
        self.MaximizeBox = False
        _apply_logo(self)

        y = 12
        if description:
            desc = Label()
            desc.Text = description
            desc.Location = Point(12, y)
            desc.Size = Size(self.Width - 40, 36)
            desc.AutoSize = False
            self.Controls.Add(desc)
            y += 42

        prefix_label = Label()
        prefix_label.Text = "Prefix"
        prefix_label.Location = Point(12, y)
        prefix_label.Size = Size(80, 20)
        self.Controls.Add(prefix_label)

        self._prefix_input = TextBox()
        self._prefix_input.Location = Point(100, y)
        self._prefix_input.Size = Size(360, 20)
        self._prefix_input.Text = prefix_default or ""
        self.Controls.Add(self._prefix_input)
        y += 32

        suffix_label = Label()
        suffix_label.Text = "Suffix"
        suffix_label.Location = Point(12, y)
        suffix_label.Size = Size(80, 20)
        self.Controls.Add(suffix_label)

        self._suffix_input = TextBox()
        self._suffix_input.Location = Point(100, y)
        self._suffix_input.Size = Size(360, 20)
        self._suffix_input.Text = suffix_default or ""
        self.Controls.Add(self._suffix_input)
        y += 36

        group = GroupBox()
        group.Text = "Duplicate Option"
        group.Location = Point(12, y)
        group.Size = Size(self.Width - 40, 110)
        self.Controls.Add(group)

        self._radio_buttons = []
        for idx, option in enumerate(duplicate_options or []):
            label, value = option
            rb = RadioButton()
            rb.Text = label
            rb.Tag = value
            rb.Location = Point(12, 24 + idx * 22)
            rb.AutoSize = True
            rb.Checked = idx == default_index
            group.Controls.Add(rb)
            self._radio_buttons.append(rb)

        ok_btn = Button()
        ok_btn.Text = button_text or "Set Values"
        ok_btn.DialogResult = DialogResult.OK
        ok_btn.Location = Point(self.Width - 220, self.Height - 80)
        self.Controls.Add(ok_btn)

        cancel_btn = Button()
        cancel_btn.Text = "Cancel"
        cancel_btn.DialogResult = DialogResult.Cancel
        cancel_btn.Location = Point(self.Width - 120, self.Height - 80)
        self.Controls.Add(cancel_btn)

        self.AcceptButton = ok_btn
        self.CancelButton = cancel_btn

    @property
    def prefix(self):
        return self._prefix_input.Text or ""

    @property
    def suffix(self):
        return self._suffix_input.Text or ""

    @property
    def option_value(self):
        for rb in self._radio_buttons:
            if rb.Checked:
                return rb.Tag
        return None


def UIform(
    title,
    description,
    prefix_default,
    suffix_default,
    duplicate_options,
    default_index=0,
    button_text="Set Values",
):
    form = _UIForm(
        title=title,
        description=description,
        prefix_default=prefix_default,
        suffix_default=suffix_default,
        duplicate_options=duplicate_options,
        default_index=default_index,
        button_text=button_text,
    )
    if form.ShowDialog() != DialogResult.OK:
        return None
    return {
        "prefix": form.prefix,
        "suffix": form.suffix,
        "duplicate_option": form.option_value,
    }
