---
title: Managing custom properties for files using SOLIDWORKS Document Manager API
caption: Custom Properties
description: Adding, removing, modifying, reading custom properties (visible and invisible) for files using SOLIDWORKS Document Manager API
labels: [custom properties]
---
SOLIDWORKS Document Manager API provides a comprehensive set of functions to manage (add, remove, modify and read) custom properties in SOLIDWORKS files.

Custom properties can be accessed for the

* File (general) via [ISwDMDocument](https://help.solidworks.com/2018/english/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr~SolidWorks.Interop.swdocumentmgr.ISwDMDocument.html) interface
* Configuration via [ISwDMConfiguration](https://help.solidworks.com/2018/english/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr~SolidWorks.Interop.swdocumentmgr.ISwDMConfiguration.html) interface
* Cut-list items via [ISwDMCutListItem2](https://help.solidworks.com/2018/english/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr~SolidWorks.Interop.swdocumentmgr.ISwDMCutListItem2.html) interface

It is possible to read properties one-by-one or extract the values in a batch.

Library allows to extract resolved values and text expressions. It is however not possible to resolve the value and only cached values can be extracted. For example if configuration specific property contains an expression which evaluates the mass of the model and the configuration was never activated, Document Manager won't be able to extract the calculated value until the model is opened and saved and configuration is activated and rebuilt.

Document manager additionally enables to manage invisible properties which are not present in the *Custom Properties* dialog and can only be read and written via Document Manager API.

Explore the articles in this section for more information and code examples.