---
layout: sw-tool
title: VBA Macro to hide all sketches in the model using SOLIDWORKS API
caption: Hide All Sketches
description: Macro will hide (blank) or show (unblank) all sketches (2D and 3D) in the active document using SOLIDWORKS API
image: /solidworks-api/document/sketch/hide-all-sketches/sw-hide-all-sketches.png
labels: [blank sketch, hide sketch, solidworks api, utils]
categories: sw-tools
group: Sketch
redirect-from:

  - /2018/03/solidworks-api-sketch-hide-all-sketches.html
---
This macro will hide (blank) or show (unblank) all sketches (2D and 3D) in the active document using SOLIDWORKS API.

{% include img.html src="sw-hide-all-sketches.png" width=320 alt="Hide sketch option in context menu" align="center" %}

Change *HIDE_ALL_SKETCHES* option to specify if sketches need to be hidden or shown.  

{% code-snippet { file-name: Macro.vba } %}
