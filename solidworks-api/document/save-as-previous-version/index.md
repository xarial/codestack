---
caption: Save As Previous Versions
title: VBA macro to save active file into the previous version of SOLIDWORKS
description: VBA macro to save the active SOLIDWORKS document into previous version with optional suffix and prefix
---

This VBA macro allows to save the active SOLIDWORKS document into previous versions of SOLIDWORKS.

User can specify the version to save to via **SW_VERSION** constant in the macro. This number is an offset of the version relative to the current version of SOLIDWORKS.

For example:

* If **-1** is specified for **SOLIDWORKS 2024**, then the file will be saved as **SOLIDWORKS 2023**
* If **-2** is specified for **SOLIDWORKS 2024** then the file will be saved in **SOLIDWORKS 2022**
* If **-1** is specified for **SOLIDWORKS 2025**, then the file will be saved as **SOLIDWORKS 2024**

User can specify suffix and prefix in the **PREFIX** and **SUFFIX** constants. Suffix will be applied to all references (in case assembly or drawing is saved)

~~~ vb
Const SW_VERSION As Integer = -1 'save into previous version

Const PREFIX As String = "" 'no prefix
Const SUFFIX As String = "_PREV" 'suffix is added to all references
~~~

{% code-snippet { file-name: Macro.vba } %}