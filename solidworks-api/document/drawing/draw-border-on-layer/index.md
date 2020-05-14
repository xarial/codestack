---
layout: article
title: Draw border of the active sheet on the specified layer
caption: Draw Border On Layer
description: VBA macro example demonstrates how to draw a border on the active drawing sheet on the specified layer considering the sheet scale
image: /solidworks-api/document/drawing/draw-border-on-layer/sheet-border.png
labels: [border,layer,scale]
---
{% include img.html src="sheet-border.png" width=350 alt="Sheet border drawn on the layer" align="center" %}

This VBA macro draws a border around the active sheet on the specified layer.

Macro considers sheet scale to calculate the correct coordinates of the border.

{% include_relative Macro.vba.codesnippet %}
