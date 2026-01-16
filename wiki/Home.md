# 4.1 |  WWP Tools Toolbar

{% include "../../../../.gitbook/includes/section_overview.md" %}

SvN Tools is our inhouse plugin with a variety of useful tools and automations for common workflows



{% include "../../../../.gitbook/includes/section_setup-instructions.md" %}

To setup The SvN tools, we must first step up the following:

1. Install Pyrevit
2. Setup Pyrevit
3. Install WWPTools

### Install pyRevit (required)

Location: Revit installer\
Screenshot:&#x20;

<figure><img src="../../../../.gitbook/assets/image (107).png" alt=""><figcaption></figcaption></figure>

Purpose: Required host for WWPTools and WWPTool extensions.

How to use:

{% stepper %}
{% step %}
Run the installer from `N:\Library\Design Software\Autodesk\Revit\! Add-Ins\pyRevit\pyRevit_4.8.12.22247_signed.exe`.
{% endstep %}

{% step %}
Accept the agreement and continue through all prompts.
{% endstep %}

{% step %}
Open Revit and click "Always Load" on the security prompt.
{% endstep %}
{% endstepper %}

This page documents setup steps and tools found under the WWPTools add-on tab.

***

### Install WWPTools MSI

Location: GitHub Releases\
Screenshot: (Add later)\
Purpose: Install the WWPTools add-in via the official MSI.

How to use:

{% stepper %}
{% step %}
Download the latest MSI from `https://github.com/WWP-Architects-Planners/WWPTools/releases/latest`.

* Current release: `WWPTools-v1.1.3.msi`
{% endstep %}

{% step %}
Run the MSI and follow the prompts.
{% endstep %}

{% step %}
Open Revit and verify the WWPTools tab is visible
{% endstep %}
{% endstepper %}

## Verify installed dynamo packages after installation

Though most of the scripts are written in python, there is still some need to use dynamo packages for some apps before they got fully transitioned.

To setup, open Revit, go to manage-dynamo. accept the terms if prompt any.

In Dynamo Preferences, packages

Verify packages are installed:

* Archi-lab.net
* bimorphNodes
* Clockwork for Dynamo 2.x
* Crumple
* Data-Shapes
* DynamoIronPython 2.7
* Elk
* Genius Loci
* Rhythm
* spring nodes
* WWP Packages


