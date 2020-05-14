---
layout: sw-tool
title: Automatically assign new file name for SOLIDWORKS files
caption: Assign New File Name
description: VBA macro to automatically assign new file name for the document based on the referenced drawing view or custom property using SOLIDWORKS API
image: /solidworks-api/document/set-new-file-name/save-as-dialog.png
labels: [new file name,auto name]
categories: sw-tools
group: Model
---
This VBA macro allows to automatically set the name for the new file based on the custom property value or drawing view referenced model using SOLIDWORKS API.

This macro will only run for the files which were never saved before.

{% include img.html src="save-as-dialog.png" width=350 alt="File Save As dialog" align="center" %}

## Configuration

Macro can be configured by changing the values of constants at the beginning of the macro

### Setting the name source

Source for the name can be set by changing the *NAME_SOURCE* constant which can take one of the following values

* DefaultDrawingViewFileName - extracts the name from the title of the referenced document of the view in the drawing
* DefaultDrawingViewCustomProperty - extracts the value from the custom property of the default view in the drawing
* CustomProperty - extracts the value from the custom property

If *DefaultDrawingViewCustomProperty* or *CustomProperty* option is used it is required to specify the name of the custom property to read value from in the *PRP_NAME* constant

~~~vb
Const NAME_SOURCE As Integer = NameSource_e.CustomProperty
Const PRP_NAME As String = "PartNo"
~~~

### Setting the title mode

There are 2 modes for the macro which can be set via *AUTO_SAVE* constant

* True - file will be automatically saved to the same folder as original model
* False - title will be assigned and pre-filled in the *Save As* dialog when manually saved

~~~vb
Const AUTO_SAVE As Boolean = True
~~~

{% include_relative Macro.vba.codesnippet %}
