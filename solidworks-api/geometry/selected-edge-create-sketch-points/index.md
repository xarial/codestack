---
layout: sw-tool
title: Create sketch points on selected edge via SOLIDWORKS API
caption: Create Sketch Points On Selected Edge
description: Macro creates specified number of sketch points on the selected edge in the 3D sketch
lang: en
image: /solidworks-api/geometry/selected-edge-create-sketch-points/sketch-points-edge.png
labels: [curve, evaluate, geometry, macro, points, solidworks api, spline, utility, vba]
categories: sw-tools
group: Sketch
redirect_from:
  - /2018/03/this-macro-creates-specified-number-of.html
---
This macro creates specified number of sketch points on the selected edge in the 3D sketch using SOLIDWORKS API.

1. Open SOLIDWORKS part
1. *(Optionally)* Open 3D Sketch to insert points to the existing sketch, otherwise new sketch will be created
1. Run the macro. Enter the number of points to generate

{% include img.html src="selected-edge.png" width=320 height=239 alt="Selected edge to create points on" align="center" %}

As the result specified number of sketch points is generated in the 3D sketch:

{% include img.html src="sketch-points-edge.png" width=320 height=204 alt="Sketch points created on the edge" align="center" %}

{% include_relative Macro.vba.codesnippet %}

Alternatively, it is possible to create points based on the curve length. The following example will create points by calculating the approximate length from curve tessellation points:

{% include_relative SplitCurveByLength.vba.codesnippet %}

or by calculating the distance based on the total curve length:

{% include_relative SplitCurveByChord.vba.codesnippet %}
