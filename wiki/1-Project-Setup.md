---
description: >-
  This tab of WWPTools is mainly focused on setting up the project and get all
  the required parameters loaded to your project.
icon: gear
---

# 1 Project Setup

<figure><img src="../../../../.gitbook/assets/image (267).png" alt=""><figcaption></figcaption></figure>

***

## Tool reference

<figure><img src="../../../../.gitbook/assets/image (268).png" alt=""><figcaption></figcaption></figure>

### Line Type Tool

Location: `WWPTools > 1. Project Setup > Line Type Tool`\
Purpose: Import and standardize line types in the project.

Use this to bulk-load office-standard line types into an existing model.\
The tool audits line types, imports missing types, and replaces mismatched names.

For one-off line types, use Revit:

`Manage > Additional Settings > Line Styles`

<figure><img src="../../../../.gitbook/assets/image (270).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../../../../.gitbook/assets/image (273).png" alt=""><figcaption></figcaption></figure>

Source file: `N:\Library\Design Software\Autodesk\Revit\Standards\Line Type\Linetypes.xlsx`\
(Planned move to GitHub in the future.)

#### How to use

{% stepper %}
{% step %}
### Run

Run the tool and follow the prompts.
{% endstep %}

{% step %}
### Review

Review the results.\
Undo if needed.
{% endstep %}
{% endstepper %}

{% hint style="info" %}
If you are using the WWP Revit Template V3+, you usually do not need this tool.\
Line styles should already be preloaded.
{% endhint %}

***

### Add Project Parameter (Pulldown)

Purpose: Import standard shared parameters, or generate an import template.

<figure><img src="../../../../.gitbook/assets/image (269).png" alt=""><figcaption></figcaption></figure>

#### Import from Excel

Location: `WWPTools > 1. Project Setup > Add Project Parameter > Import from Excel`\
Purpose: Import shared parameters from a workbook.

Source file: `N:\Library\Design Software\Autodesk\Revit\Shared Parameters\Shared Parameters Import.xlsx`

Use this tool to bulk-load office-standard shared parameters into an existing project.\
It audits project parameters, then imports any missing ones.

For one-off parameters, use Revit:

`Manage > Project Parameters`

<figure><img src="../../../../.gitbook/assets/image (274).png" alt=""><figcaption></figcaption></figure>

{% hint style="info" %}
If a project needs new standards (consultant BEP requirements, shared standards, etc.), contact the BIM team.
{% endhint %}

**How to use**

{% stepper %}
{% step %}
### Prepare

Have the workbook available.\
Select the sheet(s) when prompted.
{% endstep %}

{% step %}
### Run

Run the tool and follow the prompts.
{% endstep %}

{% step %}
### Review

Review the results.\
Undo if needed.
{% endstep %}
{% endstepper %}

<figure><img src="../../../../.gitbook/assets/image (275).png" alt=""><figcaption></figcaption></figure>

#### Create Template

Location: `WWPTools > 1. Project Setup > Add Project Parameter > Create Template`\
Purpose: Create an Excel template with shared parameters listed in the `Variables` sheet.

Use this when you do not have an import workbook.\
It creates a template file for bulk parameter import.

**How to use**

{% stepper %}
{% step %}
### Run

Run the tool and follow the prompts.
{% endstep %}

{% step %}
### Fill

Fill in the template workbook as needed.
{% endstep %}
{% endstepper %}

***

### Levels Setup (Pulldown)

<figure><img src="../../../../.gitbook/assets/image (277).png" alt="" width="355"><figcaption></figcaption></figure>

Purpose: Manage levels and generate views.

#### Levels Setup

Location: `WWPTools > 1. Project Setup > Levels Setup > Levels Setup`\
![](<../../../../.gitbook/assets/image (276).png>)\
Purpose: Add or remove levels based on a target floor count.

**How to use**

{% stepper %}
{% step %}
### Enter floors

Enter the desired number of floors.
{% endstep %}

{% step %}
### Confirm

Confirm the list of levels being added or deleted.
{% endstep %}
{% endstepper %}

#### Views Creator

<figure><img src="../../../../.gitbook/assets/image (278).png" alt="" width="355"><figcaption></figcaption></figure>

Location: `WWPTools > 1. Project Setup > Levels Setup > Views Creator`\
Screenshot:\
Purpose: Create Area Plans, Floor Plans, and RCPs for selected levels.

**How to use**

{% stepper %}
{% step %}
### Choose types

Choose categories and view types (Area Plan, Floor Plan, or RCP).

<figure><img src="../../../../.gitbook/assets/image (280).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../../../../.gitbook/assets/image (281).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Select levels

Select the floors to create.

<figure><img src="../../../../.gitbook/assets/image (282).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Create

Click "Set Value" and wait for results.

<figure><img src="../../../../.gitbook/assets/image (286).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Project Upgrader

<figure><img src="../../../../.gitbook/assets/image (283).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 1. Project Setup > Project Upgrader`\
Purpose: Batch-upgrade families or models to the current Revit version.

#### How to use

{% stepper %}
{% step %}
### Select source

Select the source folder and click Select Folder.
{% endstep %}

{% step %}
### Wait

Wait for the tool to process and rename upgraded files.
{% endstep %}
{% endstepper %}

<figure><img src="../../../../.gitbook/assets/image (285).png" alt=""><figcaption></figcaption></figure>

