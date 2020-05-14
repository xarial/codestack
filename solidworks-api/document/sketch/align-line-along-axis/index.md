---
layout: article
title: Align line along axis using SOLIDWORKS API
caption: Align Line Along Axis
description: Example demonstrates how to align all sketch lines in the active sketch (add the sketch relation) with one of the selected options (along X, Y or Z)
image: /solidworks-api/document/sketch/align-line-along-axis/sw-sketch-line-relation.png
labels: [example, horizontal, relation, sketch, solidworks api, vertical]
redirect-from:

  - /2018/03/solidworks-api-sketch-align-line-relations.html
---
Example demonstrates how to align all sketch lines in the active sketch (add the sketch relation) with one of the selected options using SOLIDWORKS API:

* Along X (horizontal)
* Along Y (vertical)
* Along Z

This example will work with both 2D and 3D sketch.

[ISketchRelationManager](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isketchrelationmanager.html) SOLIDWORKS API interface is used to manage the relations of the sketch entities.

{% include img.html src="sw-sketch-line-relation.png" width=320 height=229 alt="Relations in sketch line" align="center" %}

{% code-snippet { file-name: Macro.vba } %}
