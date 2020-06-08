---
title: Add components to assembly using SOLIDWORKS API
caption: Add Components To Assembly
description: Example Demonstrates 2 different ways to add component into the assembly tree (single component add or batch adding)
labels: [add component, assembly, example, solidworks api]
redirect-from:
  - /2018/03/solidworks-api-assembly-add-components.html
  - /solidworks-api/document/assembly/add-components
---
This example Demonstrates 2 different ways to add component into the assembly tree using SOLIDWORKS API.

* Traditional way to add component via [AddComponentX](http://help.solidworks.com/2015/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IAssemblyDoc~AddComponent5.html) SOLIDWORKS API requires the model to be loaded into the memory. Otherwise the operation fails.
* More advanced way is to use the [AddComponents ](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc~AddComponents3.html) SOLIDWORKS API. This method allows batch insertion of different components without the need to open the model beforehand.

[Download Example Files](parts.zip)

{% code-snippet { file-name: Macro.vba } %}