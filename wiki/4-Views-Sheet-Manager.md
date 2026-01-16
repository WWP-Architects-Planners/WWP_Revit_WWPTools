---
description: >-
  Reference for WWPTools utilities to duplicate/rename views, export schedules,
  renumber sheets, purge unused views, and manage sheets in bulk.
icon: table-cells-large
---

# 4 Views Sheet Manager

This page documents the tools found in this WWPTools panel.

<figure><img src="../../../../.gitbook/assets/image (346).png" alt=""><figcaption></figcaption></figure>

***

## Tool reference

### Duplicate Views

<figure><img src="../../../../.gitbook/assets/image (347).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 4. Views Sheet Manager > Duplicate Views`\
Purpose: Duplicate selected views and bulk update the new view names.

#### How to use

{% stepper %}
{% step %}
### Select views

Select the views you want to duplicate.

<figure><img src="../../../../.gitbook/assets/image (348).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Set find text (optional)

Enter the text to replace.\
Leave blank if you do not need find/replace.

<figure><img src="../../../../.gitbook/assets/image (350).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Set replacement or suffix

Enter the replacement text or the suffix to append.

<figure><img src="../../../../.gitbook/assets/image (351).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Choose duplicate mode and run

Choose the duplicate option and run the tool.

<figure><img src="../../../../.gitbook/assets/image (352).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../../../../.gitbook/assets/image (355).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Replace View Name (Bulk Views Renamer)

<figure><img src="../../../../.gitbook/assets/image (358).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 4. Views Sheet Manager > Replace View Name`\
Purpose: Find and replace text in view names, with optional prefix/suffix.

#### How to use

{% stepper %}
{% step %}
### Set find/replace

Enter the search text and replacement text.

<figure><img src="../../../../.gitbook/assets/image (377).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Add prefix/suffix (optional)

Add a prefix or suffix if needed.

<figure><img src="../../../../.gitbook/assets/image (376).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run the tool and review results

Run the tool and review the results.

<figure><img src="../../../../.gitbook/assets/image (378).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Schedules to Excel/CSV

<figure><img src="../../../../.gitbook/assets/image (364).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 4. Views Sheet Manager > Schedules to Excel/CSV`\
Purpose: Export schedules to an Excel host workbook or to CSV files.

#### How to use

{% stepper %}
{% step %}
### Select schedules

Select the schedules to export.

<figure><img src="../../../../.gitbook/assets/image (365).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Choose output

Choose either:

* an Excel host file (for Excel output), or
* a folder (for CSV output).

<figure><img src="../../../../.gitbook/assets/image (366).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run the tool and review exports

Run the tool and review the exported files.

<figure><img src="../../../../.gitbook/assets/image (367).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

{% hint style="info" %}
After your first export, the tool remembers the last output location.\
The file dialog opens to that location by default.
{% endhint %}

***

### Sheets Renumber

<figure><img src="../../../../.gitbook/assets/image (368).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 4. Views Sheet Manager > Sheets Renumber`\
Purpose: Renumber a selected set of sheets.

#### How to use

{% stepper %}
{% step %}
### Select sheets

Select the target sheets.

<figure><img src="../../../../.gitbook/assets/image (370).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run the tool and follow prompts

Run the tool and follow the prompts.

<figure><img src="../../../../.gitbook/assets/image (373).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../../../../.gitbook/assets/image (374).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Purge Views/Sheets (Pulldown)

<figure><img src="../../../../.gitbook/assets/image (375).png" alt=""><figcaption></figcaption></figure>

Purpose: Delete unused views and clean up sheets and views in bulk.

#### Clean Selected Sheet Sets & Views

Location: `WWPTools > 4. Views Sheet Manager > Purge Views/Sheets > Clean Selected Sheet Sets & Views`\
Purpose: Clean views and sheets using the selected print set.

{% stepper %}
{% step %}
### Select a sheet set and views

Select the target sheet set and any views when prompted.
{% endstep %}

{% step %}
### Run the tool

Run the tool and follow the prompts.
{% endstep %}
{% endstepper %}

#### Delete Unused Views

Location: `WWPTools > 4. Views Sheet Manager > Purge Views/Sheets > Delete Unused Views`\
Purpose: Delete views that are not placed on sheets.

{% stepper %}
{% step %}
### Run the tool

<figure><img src="../../../../.gitbook/assets/image (1).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Confirm deletion

<figure><img src="../../../../.gitbook/assets/image (2).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Sheets Manager (Pulldown)

<figure><img src="../../../../.gitbook/assets/image (379).png" alt=""><figcaption></figcaption></figure>

Purpose: Duplicate or delete sheets and the views placed on them.

#### Delete Current Selected Sheet and its Views

Location: `WWPTools > 4. Views Sheet Manager > Sheets Manager > Delete Current Selected Sheet and its Views`\
Purpose: Delete selected sheets and all views on them.

{% stepper %}
{% step %}
### Select sheets

Select the target sheets.
{% endstep %}

{% step %}
### Run the tool

Run the tool and follow the prompts.

all views in the library and the current selected sheet will be deleted.
{% endstep %}
{% endstepper %}

{% hint style="info" %}
Open sheets will not be deleted.
{% endhint %}

#### Delete Sheets and its Views from List

Location: `WWPTools > 4. Views Sheet Manager > Sheets Manager > Delete Sheets and its Views from List`\
Purpose: Delete selected sheets and all views on them.

{% stepper %}
{% step %}
### Select sheets

Select the target sheets from the list.

<figure><img src="../../../../.gitbook/assets/image (380).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run the tool

Run the tool and follow the prompts.
{% endstep %}
{% endstepper %}

{% hint style="info" %}
Open sheets will not be deleted.
{% endhint %}

#### Delete Views from Current Sheet

Location: `WWPTools > 4. Views Sheet Manager > Sheets Manager > Delete Views from Current Sheet`\
Purpose: Delete views from the current sheet and the project browser.

{% stepper %}
{% step %}
### Activate sheet

Open the sheet you want to modify.
{% endstep %}

{% step %}
### Run the tool

Run the tool and follow the prompts.
{% endstep %}
{% endstepper %}

#### Sheet Duplicator

Location: `WWPTools > 4. Views Sheet Manager > Sheets Manager > Sheet Duplicator`\
Purpose: Duplicate selected sheets from a list.

{% stepper %}
{% step %}
### Select sheets

Select the target sheets.
{% endstep %}

{% step %}
### Choose view duplication

Choose whether to duplicate views placed on each sheet.

<figure><img src="../../../../.gitbook/assets/image (382).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run the tool

Run the tool and follow the prompts.
{% endstep %}
{% endstepper %}

