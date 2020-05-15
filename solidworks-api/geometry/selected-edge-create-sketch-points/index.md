---
layout: sw-tool
title: Create sketch points on selected edge via SOLIDWORKS API
caption: Create Sketch Points On Selected Edge
description: Macro creates specified number of sketch points on the selected edge in the 3D sketch
image: sketch-points-edge.png
labels: [curve, evaluate, geometry, macro, points, solidworks api, spline, utility, vba]
group: Sketch
redirect-from:
  - /2018/03/this-macro-creates-specified-number-of.html
---
This macro creates specified number of sketch points on the selected edge in the 3D sketch using SOLIDWORKS API.

1. Open SOLIDWORKS part
1. *(Optionally)* Open 3D Sketch to insert points to the existing sketch, otherwise new sketch will be created
1. Run the macro. Enter the number of points to generate

![Selected edge to create points on](selected-edge.png){ width=320 height=239 }

As the result specified number of sketch points is generated in the 3D sketch:

![Sketch points created on the edge](sketch-points-edge.png){ width=320 height=204 }

{% code-snippet { file-name: Macro.vba } %}

Alternatively, it is possible to create points based on the curve length. The following example will create points by calculating the approximate length from curve tessellation points:

{% code-snippet { file-name: SplitCurveByLength.vba } %}

or by calculating the distance based on the total curve length:

{% code-snippet { file-name: SplitCurveByChord.vba } %}
