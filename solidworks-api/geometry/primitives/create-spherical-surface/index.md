---
title: Create temp spherical sheet body using SOLIDWORKS modeler API
caption: Create Temp Spherical Sheet Body
description: Example demonstrates how to create temp body of a spherical sheet
image: spherical-surface.png
labels: [topology, geometry, sheet, modeler, sphere]
---
![Spherical sheet body](spherical-surface.png)

This example demonstrates how to create a sheet body from the spherical surface using SOLIDWORKS API.

Geometry is created using the [IModeler::CreateSphericalSurface2](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeler~createsphericalsurface2.html) SOLIDWORKS API method.

Run the macro and temp body is displayed. Body can be rotated and selected but it is not presented in the feature tree. Continue the macro execution to destroy the body.

{% code-snippet { file-name: Macro.vba } %}
