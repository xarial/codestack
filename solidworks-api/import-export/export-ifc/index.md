---
layout: article
title: Export assembly or part to IFC 2x3 or 4 using SOLIDWORKS API
caption: Export To IFC (2x3 and 4)
description: Example in VBA to export active part or assembly document to IFC format (2x3 or 4) using SOLIDWORKS API
image: /solidworks-api/import-export/export-ifc/save-as-dialog-ifc.png
labels: [ifc,2x3,4,export]
categories: sw-tools
group: Model
---
This example demonstrates how to export active assembly or part document into IFC format using SOLIDWORKS API.

{% include img.html src="save-as-dialog-ifc.png" width=450 alt="Save As Dialog with 2 ifc format options" align="center" %}

Currently SOLIDWORKS supports 2 schemas for IFC format

* IFC 2x3
* IFC 4

The schema can be validated in the output IFC file when opened in text editor.

{% include img.html src="ifc-file.png" width=450 alt="IFC schema" align="center" %}

This VBA macro demonstrates how to export file to both IFC schemas. Change the value of the *IfcFormat_e* enumeration to change the format:

~~~ vb
ExportIfc swModel, OUT_FILE_PATH, IfcFormat_e.Ifc4 'Export to IFC 4
ExportIfc swModel, OUT_FILE_PATH, IfcFormat_e.Ifc2x3 'Export to IFC 2x3
~~~

Change the value of *OUT_FILE_PATH* constant to specify the output file location

~~~ vb
Const OUT_FILE_PATH As String = "C:\Output\Building.ifc"
~~~

{% code-snippet { file-name: Macro.vba } %}
