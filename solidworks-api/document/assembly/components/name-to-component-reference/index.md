---
layout: sw-tool
title: Copy component name to the component reference using SOLIDWORKS API
caption: Copy Component Name To Component Reference
description: VBA macro to copy component name to the component reference using SOLIDWORKS with an ability to filter virtual components only
image: /solidworks-api/document/assembly/components/name-to-component-reference/component-reference.png
labels: [name,virtual,component reference]
categories: sw-tools
group: Assembly
---
{% include img.html src="component-reference.png" width=350 alt="Component reference" align="center" %}

This VBA macro allows to copy the name of the components in the active assembly to the component's reference using SOLIDWORKS API.

Macro has an option to only process virtual components by settings the *VIRTUAL_ONLY* option to *True*.

~~~ vb
Const VIRTUAL_ONLY As Boolean = True
~~~

This macro can be useful if component names are used to store the project attributes (e.g. Part Number) as component name cannot be added to the Bill Of Materials while Component Reference can be.

{% include img.html src="bill-of-materials.png" width=350 alt="Bill Of Materials with component references" align="center" %}

{% code-snippet { file-name: Macro.vba } %}
