---
layout: sw-tool
title: Macro to copy SOLIDWORKS custom property from material to model
caption: Copy Custom Property From Material To Model
description: Macro demonstrates how to copy the specified custom property from the material database to the model custom property using SOLIDWORKS API and XML parsers
image: /solidworks-api/document/materials/copy-custom-property/material-custom-property.png
labels: [material, xml, custom property]
categories: sw-tools
group: Materials
---
![Custom property in the material](material-custom-property.png){ width=450 }

This macro demonstrates how to copy the specified custom property from the material database to the model custom property using SOLIDWORKS API and XML parsers.

*MSXML2.DOMDocument* object is used to read XML of the material database and select the required material node.

* Specify the custom property name to copy via *PRP_NAME* variable

~~~ vb
Const PRP_NAME As String = "MyProperty"
~~~

* Run the macro. Macro will find the material of active part and read the property value from the material database file
* Macro will create/update the generic custom property of the file to the corresponding value of the custom property from material

{% code-snippet { file-name: Macro.vba } %}
