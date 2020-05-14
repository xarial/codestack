---
layout: article
title: Get parameters of face at centroid using SOLIDWORKS API
caption: Get Face Center Parameters
description: Example demonstrates how to find the face parameters (coordinate and normal) at the center of the face using SOLIDWORKS API
image: /solidworks-api/geometry/get-face-center-parameters/face-center.png
labels: [center, uv, normal]
---
{% include img.html src="face-center.png" width=250 alt="Point created at the center of the face" align="center" %}

This example demonstrate how to find the parameters (point coordinate and normal) at the center of the face using SOLIDWORKS API. This macro will work with any type of face (planar, cylindrical, toroidal, b-surface etc.)

Center is found as the average of minimum and maximum values of U and V parameters using the [ISurface::Evaluate](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isurface~evaluate.html) SOLIDWORKS API method.

{% code-snippet { file-name: Macro.vba } %}