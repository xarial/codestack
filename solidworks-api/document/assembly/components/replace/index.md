---
layout: sw-tool
title: Replace components preserving selection using SOLIDWORKS API
caption: Replace Components
description: Macro demonstrates how to replace selected components in the batch preserving original selections using SOLIDWORKS API
image: replace_components.png
labels: [component, replace, selection]
group: Assembly
---
![Components replaced in the tree](replace_components.png){ width=350 }

This macro allows to replace the selected components in the tree with the components from the nominated folder (optionally with additional suffix in name) using SOLIDWORKS API.

This could be useful when managing similar types of projects where some files were copied, updated and renamed and need to be replaced in the original assembly.

This macro is using the [API only selections](solidworks-api/document/selection/api-only-selection/) which allows to keep the original selected components and avoiding the need to use the temp collection variables to satisfy the requirement of [IAssemblyDoc::ReplaceComponents2](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iassemblydoc~replacecomponents2.html) SOLIDWORKS API method where each component needs to be selected for replacement.

* Modify the input parameters. Set the directory where the replacement parts are located via *REPLACEMENT_DIR* and optional *SUFFIX* for file name.

~~~ vb
Const REPLACEMENT_DIR As String = "D:\Assembly\Replacement"
Const SUFFIX As String = "_new"
~~~

* Select components
* Run macro. All components are replaced

{% code-snippet { file-name: Macro.vba } %}
