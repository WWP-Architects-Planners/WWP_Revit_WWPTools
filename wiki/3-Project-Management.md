---
description: >-
  Reference for WWPTools project-wide utilities (parameters, type renaming,
  filter/template copying, sheet scale updates, door/fire rating tools, and
  marketing view generation).
icon: diagram-project
---

# 3 Project Management

This page documents the tools found under this WWPTools panel.

<figure><img src="../../../../.gitbook/assets/image (3).png" alt=""><figcaption></figcaption></figure>

***

## Tool reference

### Copy Color Scheme

<figure><img src="../../../../.gitbook/assets/image (1) (1).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 3. Project Management > Copy Color Scheme`\
Purpose: Copy color scheme settings from one scheme to another.

#### How to use

{% stepper %}
{% step %}
### Select schemes

Select the source scheme and the target scheme.

<figure><img src="../../../../.gitbook/assets/image (3) (1).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../../../../.gitbook/assets/image (2) (1).png" alt=""><figcaption></figcaption></figure>

<figure><img src="../../../../.gitbook/assets/image (10).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Copy

Run the tool to copy the settings.
{% endstep %}
{% endstepper %}

<figure><img src="../../../../.gitbook/assets/image (6).png" alt=""><figcaption></figcaption></figure>

***

### Copy Parameter

<figure><img src="../../../../.gitbook/assets/image (7).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 3. Project Management > Copy Parameter`\
Purpose: Copy values from one instance parameter to another, with optional find/replace.

#### How to use

{% stepper %}
{% step %}
### Select elements

Select the elements to update.

<figure><img src="../../../../.gitbook/assets/image (16).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Set source + target

Enter the source and target parameter names.

<figure><img src="../../../../.gitbook/assets/image (14).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### (Optional) Find/replace

Enter find/replace strings to transform values during the copy.
{% endstep %}

{% step %}
### Write values

Click **Set Values**.

<figure><img src="../../../../.gitbook/assets/image (15).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Parking Count in Room

<figure><img src="../../../../.gitbook/assets/image (29).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 3. Project Management > Parking Count in Room`\
Purpose: Count parking by room and populate the parking occupancy room tag.

#### How to use

{% stepper %}
{% step %}
### Prepare parking families

Fill the required parameters on the parking families.
{% endstep %}

{% step %}
### Run

Select a room and run the tool.

<figure><img src="../../../../.gitbook/assets/image (17).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Verify

Confirm the tag family `_WWP_A_RoomTag_ParkingOccupancy` is populated.
{% endstep %}
{% endstepper %}

***

### Replace Type Name

<figure><img src="../../../../.gitbook/assets/image (30).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 3. Project Management > Replace Type Name`\
Purpose: Batch rename type names using find/replace.

<figure><img src="../../../../.gitbook/assets/image (18).png" alt=""><figcaption></figcaption></figure>

#### How to use

{% stepper %}
{% step %}
### Set find/replace

Enter the search text and replacement text.

<figure><img src="../../../../.gitbook/assets/image (19).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run

Run the tool and wait for the result.

<figure><img src="../../../../.gitbook/assets/image (20).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Sheet Scale Updater

<figure><img src="../../../../.gitbook/assets/image (31).png" alt=""><figcaption></figcaption></figure>

Location: `WWPTools > 3. Project Management > Sheet Scale Updater`\
Purpose: Update the titleblock **Sheet Scale** parameter from the views placed on a sheet.

#### How to use

{% stepper %}
{% step %}
### Prepare

Make sure your titleblock contains the `WWP_ScaleBar` family.

Have the target sheet(s) open or selected when prompted.

<figure><img src="../../../../.gitbook/assets/image (21).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run

Run the tool and follow the prompts.

<figure><img src="../../../../.gitbook/assets/image (22).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

The scale bar updates on the sheet(s).

<figure><img src="../../../../.gitbook/assets/image (24).png" alt=""><figcaption></figcaption></figure>

If a sheet has multiple views with different scales, the scale bar shows **Varies**.

![](<../../../../.gitbook/assets/image (26).png>)<br>

***

### Copy Filters (Pulldown)

<figure><img src="../../../../.gitbook/assets/image (28).png" alt=""><figcaption></figcaption></figure>

Purpose: Copy filter graphic overrides between views and view templates.

<figure><img src="../../../../.gitbook/assets/image (27).png" alt=""><figcaption></figcaption></figure>

#### Copy Current Filters to Multiple Templates

Location: `WWPTools > 3. Project Management > Copy Filters > Copy Current Filters to Multiple Templates`\
Purpose: Copy filter graphic overrides from the current view to multiple templates.

**How to use**

{% stepper %}
{% step %}
### Open source view

Open the view that contains the desired filter overrides.
{% endstep %}

{% step %}
### Run

Run the tool and follow the prompts.
{% endstep %}
{% endstepper %}

#### Copy Filters to Current View

Location: `WWPTools > 3. Project Management > Copy Filters > Copy Filters to Current View`\
Purpose: Copy filter graphic overrides from a view template into the current view.

<figure><img src="../../../../.gitbook/assets/image (32).png" alt=""><figcaption></figcaption></figure>

**How to use**

{% stepper %}
{% step %}
### Open target view

Open the view you want to update.
{% endstep %}

{% step %}
### Run

Run the tool and follow the prompts.
{% endstep %}
{% endstepper %}

#### Copy Filters to Template

Location: `WWPTools > 3. Project Management > Copy Filters > Copy Filters to Template`\
Purpose: Copy filter graphic overrides from one view template to another.

**How to use**

{% stepper %}
{% step %}
### Run

Run the tool and follow the prompts.

<figure><img src="../../../../.gitbook/assets/image (33).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Door Tool (Pulldown)

<figure><img src="../../../../.gitbook/assets/image (319).png" alt=""><figcaption></figcaption></figure>

Purpose: Bulk edit door data and types.

#### Door Type Duplicator

Location: `WWPTools > 3. Project Management > Door Tool > Door Type Duplicator`\
Purpose: Duplicate a selected door type with specified dimensions.

**How to use**

{% stepper %}
{% step %}
### Prepare

Make sure door elements are available in the active view or selection.

<figure><img src="../../../../.gitbook/assets/image (320).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run

Run the tool and follow the prompts.
{% endstep %}
{% endstepper %}

#### Get Door Number from Room

Location: `WWPTools > 3. Project Management > Door Tool > Get Door Number from Room`\
Purpose: Write door numbers from **To Room** values into the door **Mark**.

<figure><img src="../../../../.gitbook/assets/image (321).png" alt=""><figcaption></figcaption></figure>

**How to use**

{% stepper %}
{% step %}
### Prepare

Make sure door elements are available in the active view or selection.
{% endstep %}

{% step %}
### Run

Run the tool and follow prompts.

Rooms with multiple doors get automatic suffixes: `104A`, `104B`, `104C`, etc.

<figure><img src="../../../../.gitbook/assets/image (324).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

{% hint style="warning" %}
This overwrites **all doors** in the project.\
It does not work for **doors inside model groups**.
{% endhint %}

#### Publish Fire Rating to Doors

Location: `WWPTools > 3. Project Management > Door Tool > Publish Fire Rating to Doors`\
Purpose: Copy wall fire rating (FRR) to a door parameter.

<figure><img src="../../../../.gitbook/assets/image (325).png" alt=""><figcaption></figcaption></figure>

**How to use**

{% stepper %}
{% step %}
### Select elements

Select the target doors and their host walls.

<figure><img src="../../../../.gitbook/assets/image (326).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Set source parameter

Enter the wall parameter name.\
Example: `FRR`.

<figure><img src="../../../../.gitbook/assets/image (327).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run

Run the tool and follow prompts.

The door parameter (example: **FRR Walls**) updates automatically and can be scheduled.
{% endstep %}
{% endstepper %}

#### Select Doors in View

Location: `WWPTools > 3. Project Management > Door Tool > Select Doors in View`\
Purpose: Select all doors in the current view.

**How to use**

{% stepper %}
{% step %}
### Open target view

Open the view you want to affect.
{% endstep %}

{% step %}
### Run

Run the tool.\
All doors in the current view will be selected.
{% endstep %}
{% endstepper %}

***

### Fire Rating Tool (Pulldown)

<figure><img src="../../../../.gitbook/assets/image (328).png" alt=""><figcaption></figcaption></figure>

Purpose: Create or clear fire rating detail lines for walls.

#### Clear All Fire Rating Lines in Current View

Location: `WWPTools > 3. Project Management > Fire Rating Tool > Clear All fire rating lines in Current View`\
Purpose: Clear all auto-generated fire rating lines in the current view.

**How to use**

{% stepper %}
{% step %}
### Open target view

Open the view you want to clean.
{% endstep %}

{% step %}
### Run

Run the tool and follow prompts.\
All fire rating lines will be cleared.

<figure><img src="../../../../.gitbook/assets/image (329).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

#### Convert Lines to Detail Items

Location: `WWPTools > 3. Project Management > Fire Rating Tool > Convert Lines to Detail Items`\
Purpose: Convert rated detail lines into line-based detail components.

Use this when you want the rating graphics as a family, not detail lines.

<figure><img src="../../../../.gitbook/assets/image (337).png" alt=""><figcaption></figcaption></figure>

**How to use**

{% stepper %}
{% step %}
### Load family

Load: `N:\Library\Design Software\Autodesk\Revit\Families\Annotations\Families\WWP_FireRating_FRR Line Based.rfa`.

<figure><img src="../../../../.gitbook/assets/image (336).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run

Run the tool and follow prompts.
{% endstep %}
{% endstepper %}

<figure><img src="../../../../.gitbook/assets/image (330).png" alt=""><figcaption></figcaption></figure>

#### Create Fire Rating Lines for ALL Walls in Current View

Location: `WWPTools > 3. Project Management > Fire Rating Tool > Create Fire Rating Lines for ALL walls in Current View`\
Purpose: Create rated lines based on wall FRR values.

**How to use**

{% stepper %}
{% step %}
### Open target view

Open the view you want to affect.

<figure><img src="../../../../.gitbook/assets/image (331).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Run

Run the tool and follow prompts.\
All walls with a value in the shared parameter `FRR` receive a fire rating detail line.

<figure><img src="../../../../.gitbook/assets/image (332).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

{% hint style="info" %}
Accepted FRR values: `0HR`, `.75HR`, `1HR`, `1.5HR`, `2HR`, `3HR`.
{% endhint %}

<figure><img src="../../../../.gitbook/assets/image (333).png" alt=""><figcaption></figcaption></figure>

#### Create Fire Rating Lines for All Walls in Views with FRR Indicator

Location: `WWPTools > 3. Project Management > Fire Rating Tool > Create Fire Rating Lines for All walls in views with FRR indicator`\
Purpose: Create fire rating lines for all views that end with `FRR`.

**How to use**

{% stepper %}
{% step %}
### Run

Run the tool and follow prompts.

<figure><img src="../../../../.gitbook/assets/image (356).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Wait

Wait for results. now all views with "FRR" will be updated with fire rating detail lines.

<figure><img src="../../../../.gitbook/assets/image (357).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

#### Create Fire Rating Lines for Selected Walls in Current View

Location: `WWPTools > 3. Project Management > Fire Rating Tool > Create Fire Rating Lines for Selected walls in Current View`\
Purpose: Create rated lines for selected walls.

**How to use**

{% stepper %}
{% step %}
### Select walls

Select walls in the current view.
{% endstep %}

{% step %}
### Run

Run the tool and follow prompts.
{% endstep %}
{% endstepper %}

#### Create Non-Rated Line for ALL Walls in Current View

Location: `WWPTools > 3. Project Management > Fire Rating Tool > Create Non-Rated Line for ALL walls in Current View`\
Purpose: Create solid centerline detail lines for all walls, regardless of FRR values.

**How to use**

{% stepper %}
{% step %}
### Open target view

Open the view you want to affect.
{% endstep %}

{% step %}
<figure><img src="../../../../.gitbook/assets/image (334).png" alt=""><figcaption></figcaption></figure>

### Run

Run the tool and follow prompts.\
All walls generate non-rated lines.\
Use this for early design stages only.

<figure><img src="../../../../.gitbook/assets/image (335).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

***

### Marketing View Maker (Pulldown)

<figure><img src="../../../../.gitbook/assets/image (338).png" alt=""><figcaption></figcaption></figure>

Purpose: Create marketing blackline views and sheets from selected areas and doors.

#### Make MULTIPLE KeyPlans from Selected Area

Location: `WWPTools > 3. Project Management > Marketing View Maker > Make MULTIPLE KeyPlans from Selected Area`\
Purpose: Create multiple keyplans based on selected areas.

**How to use**

{% stepper %}
{% step %}
### Select areas

Select one or more areas.
{% endstep %}

{% step %}
### Run

Run the tool and follow prompts.\
It creates keyplans using your templates and shades the selected area.

<figure><img src="../../../../.gitbook/assets/image (345).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

#### Make Single Suite (Recommended)

Location: `WWPTools > 3. Project Management > Marketing View Maker > Make Single Suite (Recommended)`\
Purpose: Create a suite blackline view and sheet from a selected area and entry door.

**How to use**

{% stepper %}
{% step %}
### Select inputs

Select an area and its entry door.

<figure><img src="../../../../.gitbook/assets/image (339).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Choose templates

Run the tool and follow prompts.\
Select your suite view template, keyplan template, and titleblock family.

<figure><img src="../../../../.gitbook/assets/image (340).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Auto-fill

The tool auto-fills data based on the area properties.

<figure><img src="../../../../.gitbook/assets/image (341).png" alt=""><figcaption></figcaption></figure>
{% endstep %}

{% step %}
### Generate sheet

The tool generates the sheet based on the selected door direction.

<figure><img src="../../../../.gitbook/assets/image (343).png" alt=""><figcaption></figcaption></figure>

With templates set correctly, the view is auto-cropped and rotated so the entrance door faces down.\
It also reads the area and fills the sheet property for **Unit Area**.\
The Imperial value is the converted version.

<figure><img src="../../../../.gitbook/assets/image (344).png" alt=""><figcaption></figcaption></figure>
{% endstep %}
{% endstepper %}

Notes:

* Creates backup sheets with `- bak` suffix if sheets already exist.
* Auto-crops and rotates to align the door direction.

