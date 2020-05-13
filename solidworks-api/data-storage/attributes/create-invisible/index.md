---
layout: article
title: Create invisible attribute using SOLIDWORKS API
caption: Create Invisible Attribute
description: Example creates an invisible attribute and attaches to the selected object (entity or component)
lang: en
image: /solidworks-api/data-storage/attributes/create-invisible/sw-attribute-features-tree.png
labels: [attribute, data, example]
redirect_from:
  - /2018/03/create-invisible-attribute.html
---
This example creates an invisible attribute and attaches to the selected object (entity or component).

Attribute ca be hidden by setting the corresponding flag in the [IAttributeDef::CreateInstance5](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iattributedef~createinstance5.html) SOLIDWORKS API method.

{% include img.html src="sw-attribute-features-tree.png" width=272 height=320 alt="Attribute feature inserted to the Feature Manager Tree" align="center" %}

Macro stops the execution once the attribute is created. At this stage the attribute feature is invisible.
When execution of macro continues (F5 or run is clicked) the feature is set to visible.

{% include_relative Macro.vba.codesnippet %}
