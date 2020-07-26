---
title: Assembly components automation using SOLIDWORKS API
caption: Components
description: Collection of articles and code examples for working with components in SOLIDWORKS assembly
labels: [assembly, components]
order: 1
---
Component in SOLIDWORKS assembly is an instance of the model documents ([IModelDoc2](https://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2.html)) in assembly.

Components can be automated via [IComponent2](https://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IComponent2.html) interface available in SOLIDWORKS API.

The main operations on components include, but not limited to:

* Transformation
* Mating
* In context editing
* BOM composition

Pointer to the underlying document of the component can be retrieved via [IComponent2::GetModelDoc2](https://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.icomponent2~getmodeldoc2.html) SOLIDWORKS API method. This method returns null for suppressed or lightweight components. Refer the [Get Model Doc For Lightweight Component](/solidworks-api/document/assembly/components/lightweight-get-model-doc/) for code example demonstrating how to retrieve the pointer to all types of components.

Explore this section for code examples and macros of automating assemblies and components.