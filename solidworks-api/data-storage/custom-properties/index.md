---
title: Managing Custom Properties using SOLIDWORKS API
caption: Custom Properties
description: Managing model, configuration and feature specific custom properties using SOLIDWORKS API
labels: [custom properties, configuration properties]
---
This section contains macros and code examples for utilizing the custom properties in SOLIDWORKS API.

Custom properties is a collection of key value pairs used in SOLIDWORKS to store the metadata. Custom properties can be associated with the model itself, its configurations or cut-list features (such as weldment or sheet metal).

Custom properties are managed via [ICustomPropertyManager](https://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICustomPropertyManager.html) SOLIDWORKS API interface.

In many cases when it is required to read the value of the custom property (e.g. for using in the file name, for exporting, etc.) property will be firstly read from the referenced configuration and if missing from the file property. This is similar to the way properties are used to populate Bill Of Materials table.

Code below demonstrates how to implement this practice in the code.

{% code-snippet { file-name: GetPropertyValue.vba } %}