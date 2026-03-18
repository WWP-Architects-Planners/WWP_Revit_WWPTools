# WWP Manual Revisions — PyRevit Tool

## What it does
Toggles between Revit's native revision schedule and a manual 4-parameter
revision block on one or more sheets. Provides a WPF dialog to fill in
revision content directly.

---

## Required Shared Parameters (on ViewSheet)

| Parameter Name                  | Type    | Notes                          |
|---------------------------------|---------|--------------------------------|
| WWP_ShowManualRevisions_YesNo   | Yes/No  | Drives visibility formulas     |
| WWP_Rev_Left_Dates              | Text    | One date per line, left column |
| WWP_Rev_Left_Descs              | Text    | One desc per line, left column |
| WWP_Rev_Right_Dates             | Text    | One date per line, right col   |
| WWP_Rev_Right_Descs             | Text    | One desc per line, right col   |

All five must be added to your Shared Parameter file and bound to sheets
before running the tool.

---

## Titleblock Setup

1. In the titleblock family, add all 5 parameters as **instance** parameters.
2. Place **Label** elements for each of the 4 text parameters inside the
   manual revision table box.
3. Set the **visibility** of the manual table group:
   - Visible when: `WWP_ShowManualRevisions_YesNo == Yes`
4. Set the **visibility** of the native Revit revision schedule:
   - Visible when: `NOT(WWP_ShowManualRevisions_YesNo)`

---

## Usage

- Open a sheet (or select multiple sheets in the Project Browser)
- Run the tool from the WWP tab
- Check/uncheck the toggle and fill in the four text fields
- One entry per line — date lines and description lines must match in count
- Click **Apply to Sheet**

For multiple sheets: all selected sheets get the same values written.

---

## Installation

Drop the `WWP_ManualRevisions.pushbutton` folder into your PyRevit
extension's panel folder, e.g.:

```
WWP.extension/
  WWP.tab/
    Sheets.panel/
      WWP_ManualRevisions.pushbutton/
        script.py
        bundle.yaml
        icon.png   ← optional 32x32 PNG
```

Reload PyRevit.
