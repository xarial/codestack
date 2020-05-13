---
layout: article
title: Get sketch lines of sheet metal sketched bend using SOLIDWORKS API
caption: Get Sketch Lines For Sheet Metal Sketched Bend Feature
description: Finds all straight lines (bends) of the sheet metal Sketched Bend feature and selects all segments
lang: en
image: /solidworks-api/document/sheet-metal/sketched-bend-feature-sketch-lines/sheet-metal-sketched-bend.png
labels: [example, sheet metal, sketched bend, solidworks api]
redirect_from:
  - /2018/03/solidworks-api-sheet-metal-get-sketched-bends.html
---
Macro finds all straight lines (bends) of the sheet metal *Sketched Bend* feature and selects all segments using SOLIDWORKS API.

{% include img.html src="sheet-metal-sketched-bend.png" width=400 alt="Sketch of the sheet metal sketched bend feature" align="center" %}

There is no direct SOLIDWORKS API method of getting the bends, however bends are represented as sketch segments in the sketch owned by sheet metal feature. So in order to find bends it is required to find this sketch and parse its content.

{% include_relative Macro.vba.codesnippet %}
