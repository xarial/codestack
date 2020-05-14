---
layout: article
title: Add smart dimension between two segments using SOLIDWORKS API
caption: Add Smart Dimension Between Two Segments
description: Example adds the dimension between 2 selected sketch segments
image: /solidworks-api/document/dimensions/add-smart-dimension-between-two-segments/dimension-name.png
labels: [dimension, example, solidworks api]
redirect-from:
  - /2018/03/solidworks-api-dimensions-add-dimensions-to-sketch-segment.html
---
This example adds the dimension between 2 selected sketch segments (e.g. sketch lines) using SOLIDWORKS API. The dimension will be placed in the middle of 2 selection points.  

{% include img.html src="dimension-name.png" width=320 height=237 alt="Dimension with name" align="center" %}

When adding dimensions programmatically using SOLIDWORKS API it is important to disable the Input Dimension Value option otherwise the macro will be interrupted and will require user inputs.

The example below temporarily removes this option and restores the original value after the dimension inserted so user settings are not affected.  

{% include img.html src="input-dimension-value-option.png" width=640 height=198 alt="Option to input dimension value on creation" align="center" %}

{% code-snippet { file-name: Macro.vba } %}
