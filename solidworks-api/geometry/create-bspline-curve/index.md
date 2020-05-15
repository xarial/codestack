---
layout: article
title: Create and display b-spline curve using SOLIDWORKS API
caption: Create B-Spline Curve
description: VBA example demonstrates how to create and preview b-spline curve from the sample data using SOLIDWORKS API
image: bspline-curve-preview.png
labels: [curve, bspline, modeler]
---
![Preview of b-spline curve](bspline-curve-preview.png){ width=250 }

This VBA example demonstrates the use of [IModeler::CreateBsplineCurve](https://help.solidworks.com/2012/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModeler~CreateBsplineCurve.html) method to create and preview b-spline curve using sample data.

Open part document and run the macro. Curve will be previewed and macro stops. Continue the macro to dispose the curve.

Follow the [Get B-Spline Curve Parameters](/solidworks-api/geometry/get-bspline-parameters/) example for a guide of extracting the required data from the selected edge.

{% code-snippet { file-name: Macro.vba } %}