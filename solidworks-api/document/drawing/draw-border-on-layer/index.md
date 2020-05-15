---
layout: article
title: Draw border of the active sheet on the specified layer
caption: Draw Border On Layer
description: VBA macro example demonstrates how to draw a border on the active drawing sheet on the specified layer considering the sheet scale
image: sheet-border.png
labels: [border,layer,scale]
---
![Sheet border drawn on the layer](sheet-border.png){ width=350 }

This VBA macro draws a border around the active sheet on the specified layer.

Macro considers sheet scale to calculate the correct coordinates of the border.

{% code-snippet { file-name: Macro.vba } %}
