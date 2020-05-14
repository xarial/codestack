---
layout: sw-tool
title: Copy custom properties from the drawing view to SOLIDWORKS drawing file
caption: Copy Drawing View Properties
description: VBA macro to copy specified custom properties from the selected or default drawing view into the drawing properties
image: /solidworks-api/document/drawing/copy-view-properties/drawing-custom-properties.png
labels: [drawing,view,custom properties]
group: Drawing
---
![Custom properties in SOLIDWORKS drawing](drawing-custom-properties.png){ width=500 }

This macro copies the specified custom properties from the SOLIDWORKS part or assembly referenced in the drawing view to the drawing view itself.

Custom properties can be specified in the *PRP_NAMES* constant in the macro. Use comma to specify multiple properties to copy.

~~~ vb
Const PRP_NAMES As String = "PartNo,Description,Title"
~~~

In order to select the properties to copy at runtime, specify an empty string as the value of *PRP_NAMES*

~~~ vb
Const PRP_NAMES As String = ""
~~~

In this case the following input box will be displayed.

![Input box for properties to be copied to drawing](properties-input-box.png)

User can specify either single property or multiple properties, separated by comma.

If drawing view is selected when running the macro, properties will be copied from this drawing view. Otherwise, the default properties view will be used as specified in the sheet properties (this is usually the first view in the drawing):

![Drawing View for custom properties](properties-view.png){ width=500 }

At first, custom property value will be extracted from the configuration of the model which corresponds to the referenced configuration of the drawing view. If the property doesn't exist or empty, file specific custom property will be extracted.

{% code-snippet { file-name: Macro.vba } %}
