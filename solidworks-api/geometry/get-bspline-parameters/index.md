---
layout: article
title: Get b-spline parameters from the selected edge using SOLIDWORKS API
caption: Get B-Spline Parameters
description: Get parameters of b-spline curve (dimension, order, periodicity, control and knot points) from the edge selected in the graphics view using SOLIDWORKS API
image: /solidworks-api/geometry/get-bspline-parameters/selected-bspline-edge.png
labels: [bspline, parameters, modeler, edge]
---
{% include img.html src="selected-bspline-edge.png" width=250 alt="Selected b-spline edge" align="center" %}

This VBA example extracts the parameters (dimension, order, periodicity, control and knot points) from the selected edge of b-spline type (e.g. edge derived from the spline segment). The extracted data can be used in the [IModeler::CreateBsplineCurve](https://help.solidworks.com/2012/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModeler~CreateBsplineCurve.html) SOLIDWORKS API method to build the curve of the same geometry.

The data is output into the Immediate window of VBA editor in the following format:

~~~
Props:
 Dimension Val 
 Order Val
 Control Points Count Val
 Periodic Val
Knots:
 Val 1
 ...
 Val N
Control Points:
 Val 1
 ...
 Val N
~~~

{% code-snippet { file-name: Macro.vba } %}