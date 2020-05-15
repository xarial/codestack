---
layout: sw-tool
title: Copy component name to the component reference using SOLIDWORKS API
caption: Copy Component Name To Component Reference
description: VBA macro to copy component name to the component reference using SOLIDWORKS with an ability to filter virtual components only
image: component-reference.png
labels: [name,virtual,component reference]
group: Assembly
---
![Component reference](component-reference.png){ width=350 }

This VBA macro allows to copy the name of the components in the active assembly to the component's reference using SOLIDWORKS API.

Macro has an option to only process virtual components by settings the *VIRTUAL_ONLY* option to *True*.

~~~ vb
Const VIRTUAL_ONLY As Boolean = True
~~~

This macro can be useful if component names are used to store the project attributes (e.g. Part Number) as component name cannot be added to the Bill Of Materials while Component Reference can be.

![Bill Of Materials with component references](bill-of-materials.png){ width=350 }

{% code-snippet { file-name: Macro.vba } %}
