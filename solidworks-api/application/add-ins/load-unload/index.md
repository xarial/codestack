---
layout: sw-tool
title: Macro to load and unload add-in using SOLIDWORKS API
caption: Load/Unload Add-In
description: Macro to trigger (load/unload) the specified add-in using SOLIDWORKS API
image: /solidworks-api/application/add-ins/load-unload/toggle-addins.png
logo: /solidworks-api/application/add-ins/load-unload/toggle-addins.svg
labels: [add-in, load]
categories: sw-tools
group: Frame
---
This macro allows to trigger the load state of the specified add-in using the [ISldWorks::LoadAddIn](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~loadaddin.html) and [ISldWorks::UnloadAddIn](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~unloadaddin.html) SOLIDWORKS API.

This can be useful to provide a short-cut for loading and unloading the add-in with one button click. It is recommended to use [Macro Buttons]({{ "solidworks-api/getting-started/macros/macro-buttons/" | relative_url }}) to create a button for add-in in the toolbar.

Macro requires the add-in Global Unique Identifier (GUID) to be specified at the beginning of the macro.

~~~ vb
Const ADD_IN_GUID As String = "{1730410d-85ad-4be8-aa2d-ed977b93fe5d}"
~~~

Locate the guid of the required SOLIDWORKS add-in in the registry at *HKLM\SOFTWARE\SolidWorks\AddIns*. Each sub-key of this registry key represents the add-in. Select each key to see the title and description of the add-in. Copy the name of the key which represents the add-in guid.

{% include img.html src="addins-registry.png" width=450 alt="Available add-ins presented in the registry" align="center" %}

It is optionally required to specify the path to the add-in in the *ADD_IN_PATH* variable. In some cases macro cannot retrieve the path to the add-in from its guid and can fail. You can find the path to the add-in in the SOLIDWORKS add-ins dialog:

{% include img.html src="addins-list.png" width=450 alt="Add-ins list in SOLIWORKS menu" align="center" %}

~~~ vb
Const ADD_IN_PATH As String = "C:\Program Files\CodeStack\MyToolbar\CodeStack.Sw.MyToolbar.dll"
~~~

If this option is not used set the value to an empty string

~~~ vb
Const ADD_IN_PATH As String = ""
~~~

{% code-snippet { file-name: Macro.vba } %}
