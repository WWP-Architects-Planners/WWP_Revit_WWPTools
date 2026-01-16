---
description: >-
  Reference for WWPTools mass/context utilities (bulk material renaming,
  CAD-to-detail line conversion, tree randomization, mass data syncing, and
  OSM/CAD context building tools).
icon: city
---

# 2 Mass Context

This page documents the tools found under this WWPTools panel.

<figure><img src="../../../../.gitbook/assets/image (287).png" alt=""><figcaption></figcaption></figure>

***

## Tool reference

### Bulk Rename Materials

<figure><img src="../../../../.gitbook/assets/image (288).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 2. Mass Context > Bulk Rename Materials`\
Screenshot: (Add later)\
Purpose: Find/replace text across all material names in the current model.

Use this to standardize naming across materials.\
Undo immediately if the result is not expected.

### How to use

{% stepper %}
{% step %}
### Run

Run the tool and follow the prompts.

Example: original material names\
![](<../../../../.gitbook/assets/image (289).png>)
{% endstep %}

{% step %}
### Find / replace

Enter the **Find** text, then the **Replace** text.

<figure><img src="../../../../.gitbook/assets/image (290).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Review and finish

Review the preview.\
Undo if needed.

<figure><img src="../../../../.gitbook/assets/image (291).png" alt=""><figcaption></figcaption></figure>

Click **Proceed**, then close the pop-up.

<figure><img src="../../../../.gitbook/assets/image (292).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### CAD Line Tool

<figure><img src="../../../../.gitbook/assets/image (293).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 2. Mass Context > CAD Line Tool`\
Purpose: Convert linked CAD layers into detail line groups in the current view.

### How to use

{% stepper %}
{% step %}
### Select CAD link

Select the linked CAD file in the view.

<figure><img src="../../../../.gitbook/assets/image (294).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Choose layer

Select the desired layer.

<figure><img src="../../../../.gitbook/assets/image (295).png" alt=""><figcaption></figcaption></figure>

Click **OK** to generate grouped detail lines by layer.
{% endstep %}

{% step %}
### Assign line types

Pre-assign line types (recommended).

<figure><img src="../../../../.gitbook/assets/image (296).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Wait

After you click **OK**, wait for the tool to finish.

<figure><img src="../../../../.gitbook/assets/image (297).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Verify

A detail group is created for the layer name.

<figure><img src="../../../../.gitbook/assets/image (298).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Random Tree

<figure><img src="../../../../.gitbook/assets/image (302).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 2. Mass Context > Random Tree`\
Purpose: Randomize tree family rotation and scale within a small range.

### How to use

{% stepper %}
{% step %}
### Select trees

Select the trees you want to update.

<figure><img src="../../../../.gitbook/assets/image (299).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Randomize

Run the tool to randomize direction and size (within 10%).

<figure><img src="../../../../.gitbook/assets/image (300).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Review

Check the result.

<figure><img src="../../../../.gitbook/assets/image (301).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Sync Mass Tool

<figure><img src="../../../../.gitbook/assets/image (303).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 2. Mass Context > Sync Mass Tool`\
Purpose: Push mass data to mass floors for GCA/GFA and FSI reporting.

### How to use

{% stepper %}
{% step %}
### Prepare

Create a mass and mass levels.
{% endstep %}

{% step %}
### Enter data

Enter data per mass.

<figure><img src="../../../../.gitbook/assets/image (304).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Sync

Run Sync Mass Tool to transfer values to mass floors.

<figure><img src="../../../../.gitbook/assets/image (305).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Context Builder (Pulldown)

<figure><img src="../../../../.gitbook/assets/image (306).png" alt=""><figcaption></figcaption></figure>

Purpose: Create context buildings and site elements from OpenStreetMap (OSM) or CAD.

Use this when there is no clean context/PDM file, or the context needs heavy revision.

Includes:

* **Building Importer** (from OSM)
* **CADBuilder** (from CAD)
* **Roads Importer** (from OSM)

#### Building Importer

Location: `WWPTools > 2. Mass Context > Context Builder > Building Importer`\
Screenshot: (Add later)\
Purpose: Create context buildings from an OpenStreetMap (OSM) file.

**How to use**

{% stepper %}
{% step %}
### Export OSM

Export an OSM file from OpenStreetMap.

<figure><img src="../../../../.gitbook/assets/image (307).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Create family

Create a new Conceptual Mass family in Revit.

<figure><img src="../../../../.gitbook/assets/image (308).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Import

Select the OSM file and units (mm or meters).

<figure><img src="../../../../.gitbook/assets/image (310).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Load

Load into the project when complete.

<figure><img src="../../../../.gitbook/assets/image (311).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

#### CADBuilder

Location: `WWPTools > 2. Mass Context > Context Builder > CADBuilder`\
Purpose: Create context buildings from a CAD file.

**How to use**

{% stepper %}
{% step %}
### Select

Select the CAD file and the target layer.

<figure><img src="../../../../.gitbook/assets/image (312).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Settings

Keep default settings unless the project needs changes.

<figure><img src="../../../../.gitbook/assets/image (313).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run

Run CADBuilder and load into the project when complete.

<figure><img src="../../../../.gitbook/assets/image (314).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

#### CADBuilder2023

Location: `WWPTools > 2. Mass Context > Context Builder > CADBuilder2023`\
Screenshot: (Add later)\
Purpose: CADBuilder workflow for Revit 2023+.

**How to use**

{% stepper %}
{% step %}
### Select

Select the CAD file and the target layer.
{% endstep %}

{% step %}
### Settings

Keep default settings unless the project needs changes.
{% endstep %}

{% step %}
### Run

Run CADBuilder2023 and load into the project when complete.
{% endstep %}
{% endstepper %}

***

#### Roads Importer

Location: `WWPTools > 2. Mass Context > Context Builder > Roads Importer`\
Screenshot: (Add later)\
Purpose: Create roads and parks from an OSM file.

**How to use**

{% stepper %}
{% step %}
### Select

Select the OSM file and units (mm or meters).
{% endstep %}

{% step %}
### Widths

Keep default widths if you are not sure.

<figure><img src="../../../../.gitbook/assets/image (315).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Load

Load into the project when complete.

<figure><img src="../../../../.gitbook/assets/image (318).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

{% hint style="warning" %}
Avoid **Copy to In-Place-Mass** unless you need editable mass in the model.\
It can slow the project down.
{% endhint %}

