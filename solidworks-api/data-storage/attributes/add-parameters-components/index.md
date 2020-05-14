---
layout: article
title: Add and read attributes with parameters to components using SOLIDWORKS API
caption: Add Attributes With Parameters To Components And Read Values
description: Example adds attributes with string values as the parameters to the selected components
image: /solidworks-api/data-storage/attributes/add-parameters-components/two-attributes-features-tree.png
labels: [attributes, data, definition, example, instance, properties, storage]
redirect-from:

  - /2018/03/add-attributes-with-parameters-to.html
---
This example adds attributes with string values as the parameters to the selected components via [IAttributeDef](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iattributedef.html) SOLIDWORKS API interface. Rebuilds the model and reads the attributes back by finding them with [IComponent2::FindAttribute](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IComponent2~FindAttribute.html) method.

Attributes are lightweight features which can be attached to SOLIDWORKS entities and store the custom data.

{% include img.html src="two-attributes-features-tree.png" width=301 height=320 alt="Two attributes features created in the Feature Manager Tree using SOLIDWORKS API" align="center" %}

{% code-snippet { file-name: Macro.vba } %}
