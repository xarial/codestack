---
layout: article
title: Selecting SOLIDWORKS Objects for API only
caption: Selecting Objects For API Only
description: Example shows how to select the object for API purpose only (without graphics selection) preserving current user selections
image: /solidworks-api/document/selection/api-only-selection/extrude-direction-up-to-surface.png
labels: [selection, extrude]
---
{% include img.html src="extrude-direction-up-to-surface.png" width=500 alt="Extruded sketch arc up to the planar surface following the line direction" align="center" %}

This example shows how to create extrude feature in SOLIDWORKS part by selecting the inputs for API purpose only (without graphics selection) preserving current user selections.

To run the macro

* Download the example file and open it in SOLIDWORKS [Extrude Selection Example](extrude-selection-example.SLDPRT)
* Select any objects (e.g. Front and Right plane)
* Debug the macro step-by-step. The macro pre-selects the required objects for the extrude feature directly in the data base (i.e. it is not visible for the user)

As the result the extrude is created with the specified direction up to specified surface and all the original user selections are preserved.

{% include_relative Macro.vba.codesnippet %}
