---
caption: Set To Read-Only
title: Macro to change read-only state of all dimensions of the selected feature in the SOLIDWORKS model
description: VBA macro to change the read-only options for all dimensions of the selected feature of the active SOLIdWORKS model
image: dimension-read-only.png
---

![Dimension read-only property](dimension-read-only.png){ width=400 }

This SOLIDWORKS VBA macro changes the read-only state of all dimensions of the selected feature (e.g. sketch).

Set the target read-only state in the constant

~~~ vb
Const READ_ONLY As Boolean = True 'True to set to Read-Only, False to remove Rea-Only flag
~~~

{% code-snippet { file-name: Macro.vba } %}