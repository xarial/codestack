---
layout: article
title: Convert arc to circle by merging end points using SOLIDWORKS API
caption: Convert Arc To Circle
description: VBA macro to convert sketch arc to a sketch circle by adding the merge relation between start and end points using SOLIDWORKS API
lang: en
image: /solidworks-api/document/sketch/convert-arc-to-circle/sketch-arc.png
labels: [sketch,arc,circle,merge,relation]
---
{% include img.html src="sketch-arc.png" width=350 alt="Sketch arc" align="center" %}

This VBA macro example demonstrates how to apply the merge sketch relation between start and end points of the selected sketch arc to convert it to sketch circle. This is the analogue of dragging the point manually until it is merged or adding the merge sketch relation in relation manager.

{% include_relative Macro.vba.codesnippet %}