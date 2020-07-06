---
layout: sw-tool
title: VBA Macro to hide all sketches in the model using SOLIDWORKS API
caption: Hide All Sketches
description: Macro will hide (blank) or show (unblank) all sketches (2D and 3D) in the active document using SOLIDWORKS API
image: hidden-sketches.svg
labels: [blank sketch, hide sketch, solidworks api, utils]
group: Sketch
redirect-from:
  - /2018/03/solidworks-api-sketch-hide-all-sketches.html
---
This macro will hide (blank) or show (unblank) all sketches (2D and 3D) in the active document using SOLIDWORKS API.

![Hide sketch option in context menu](sw-hide-all-sketches.png){ width=320 }

Change *HIDE_ALL_SKETCHES* option to specify if sketches need to be hidden or shown.

Watch [video demonstration](https://youtu.be/jsjN8zNRTuc?t=23)

{% code-snippet { file-name: Macro.vba } %}
