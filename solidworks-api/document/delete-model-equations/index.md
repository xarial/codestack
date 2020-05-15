---
layout: sw-tool
title: Delete all equations from SOLIDWORKS model using SOLIDWORKS API
caption: Delete All Equations
description: Macro removes all of the equations (or optionally only broken equations) in the active model (part or assembly)
image: /solidworks-api/document/delete-model-equations/deleted-equations.svg
labels: [api, clean, delete equations, equation, macro, utility, vba]
group: Model
redirect-from:
  - /2018/03/delete-all-equations-from-solidworks.html
---
This macro removes all of the equations (or optionally only broken equations) in the active model (part or assembly) using SOLIDWORKS API.

![Equations Manager dialog](equations-manager.png){ width=640 }

If active model is assembly, macro optionally allows to remove all equations from each component of the assembly. The following message will be displayed. Click **Yes** to remove equations from all components on all levels and **No** to only process equations of the top level assembly.

![Macro option to delete equations in the assembly components](delete-comps.png){ width=320 height=120 }

Set the *DELETE_BROKEN_ONLY* option to *True* in order to only remove the broken (dangling) equations.

**IMPORTANT: Use this macro on your own risk. This macro modifies your data (deletes all equations) please backup your file before running this macro**

{% code-snippet { file-name: Macro.vba } %}
