---
layout: article
title: Pan model views with screen pixels using SOLIDWORKS API
caption: Pan Model View
description: Example demonstrates how to pan a model view with view transforms by providing the offset in the screen pixels
lang: en
image: /solidworks-api/document/model-view/transform-pan-view/pan-view.png.png
---
{% include img.html src="pan-view.png" width=350 alt="Model View Panning" align="center" %}

This example demonstrates how to move the view (pan) by specifying the offset in X and Y coordinates of the screen (pixels). Macro transforms the offset into the model view 3D space and updates the view positions.

{% include_relative Macro.vba.codesnippet %}
