---
layout: sw-tool
title: Rename SOLIDWORKS drawing sheets with custom properties values
caption: Rename Drawing Sheets With Custom Properties Values
description: Macro will rename all drawings sheets using the value of the specified custom property using SOLIDWORKS API
image: /solidworks-api/document/drawing/rename-sheets-custom-properties-values/drw-sheets.png
labels: [custom property, drawing, example, macro, properties, rename, sheet, solidworks api, vba]
categories: sw-tools
group: Drawing
redirect_from:
  - /2018/03/document_8.html
---
This macro will rename all drawings sheets using the value of the specified custom property using SOLIDWORKS API.

{% include img.html src="drw-sheets.png" width=320 alt="List of sheets in the drawing" align="center" %}

* Open the drawing and run the macro
* Specify the property to read the value from

{% include img.html src="get-prp-name.png" width=320 alt="Popup form for property name input" align="center" %}

* All sheets are renamed based on the value of this property. Macro will get the value from the model view specified in the Sheet Properties.
The 'Same as sheet specified in Document Properties'  option is not supported.
If this option is selected then the property from the first view will be used.
Macro will try to read the configuration specific property and if the property is not specified then model level property is read.

{% include img.html src="use-custom-prps-from-view-sheet-property.png" width=400 alt="Use custom properties value from model option in the sheet properties" align="center" %}

{% code-snippet { file-name: Macro.vba } %}
