---
layout: article
title: Create temp toroidal sheet body using SOLIDWORKS modeler API
caption: Create Temp Toroidal Sheet Body
description: Example demonstrates how to create temp body of a toroidal sheet
lang: en
image: /solidworks-api/geometry/primitives/create-toroidal-surface/toroidal-surface.png
labels: [topology, geometry, sheet, modeler, cylinder]
---
{% include img.html src="toroidal-surface.png" alt="Toroidal sheet body" align="center" %}

This example demonstrates how to create a sheet body from the toroidal surface using SOLIDWORKS API.

Geometry is created using the [IModeler::CreateToroidalSurface](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeler~createtoroidalsurface.html) SOLIDWORKS API method.

Run the macro and temp body is displayed. Body can be rotated and selected but it is not presented in the feature tree. Continue the macro execution to destroy the body.

{% include_relative Macro.vba.codesnippet %}
