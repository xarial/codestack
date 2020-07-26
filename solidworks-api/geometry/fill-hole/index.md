---
title: Fill hole with temp body using SOLIDWORKS API
caption: Fill Hole
description: VBA example demonstrates how to use SOLIDWORKS modeler and create temp body to fill hole in the geometry
image: filled-hole.png
labels: [fill,modeler,hole,temp geometry]
---
![Hole filled with a temp geometry](filled-hole.png)

This VBA example demonstrates how to use [IModeler::CreateBodyFromFaces2](https://help.solidworks.com/2017/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModeler~CreateBodyFromFaces2.html) API to fill the hole of the selected feature (e.g. cut-extrude) with temp geometry.

Macro stops execution and displays temp body. Continue execution to remove the temp body.

{% code-snippet { file-name: Macro.vba } %}
