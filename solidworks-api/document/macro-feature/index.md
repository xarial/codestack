---
layout: article
title: Utilizing Macro Features in SOLIDWORKS API
caption: Macro Feature
description: Explanation and examples of using macro feature (custom feature) using SOLIDWORKS API
order: 14
---
{% youtube { id: tE_OVE9YTMs } %}

Macro feature is a type of feature which can be configured via SOLIDWORKS API and can provides same level of functionality as any native SOLIDWORKS feature.

Macro feature is inserted into the Feature Manager Tree and can be moved, deleted, suppressed or edited.

Macro feature can be inserted via [IFeatureManager::InsertMacroFeature3](http://help.solidworks.com/2014/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IFeatureManager~InsertMacroFeature3.html) method.

Macro feature definition is defined in [IMacroFeatureData](http://help.solidworks.com/2014/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IMacroFeatureData.html) SOLIDWORKS API interface

## Types Of Macro Feature

There are 2 main types of macro features: macro based and COM based. Both of this options provide the same level of functionality and only diff by supported programming language and the way they deployed and maintained.

### Macro based macro feature

This macro feature can only be created from VBA macros.

#### Benefits
* Macro can be fully embedded into the model which allows for feature to operate on any machine without the need of running any macros or installing any add-ins.
* Macro feature can be fully defined within the macro module so no need for any additional software to be registered

#### Limitations
* Maintainability. Embedded macros source code cannot be updated unless feature is deleted. However this option can be disabled so the code is centralized

### COM based macro feature

This macro feature can be created via COM-compatible language (C++, C#, VB.NET) by registering the COM server which is responsible for handling the macro feature functionality.

#### Benefits
* Centralized source code in the COM object. Simple maintenance and update

#### Limitations
* Requires the registration of the COM object on all workstations which utilizing macro feature.

## Functionality

* Execution of custom code on feature rebuild
    * On demand rebuild (ctrl+Q or ctrl+B)
    * Automatic rebuild
* Generation or modification of solid and surface bodies including the patterns
* Adding dimensions
* Storing the custom parameters within the macro feature definition
* Relationship with another entities
* Support of in-context editing in assemblies
* Support of modifications editing
* Support of custom errors
