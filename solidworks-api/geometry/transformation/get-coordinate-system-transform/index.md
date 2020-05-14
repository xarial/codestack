---
layout: article
title: Get the transformation matrix of coordinate system using SOLIDWORKS API
caption: Get Coordinate System Transformation
description: VBA macro to get the 4x4 transformation matrix from the selected coordinate systems and output the result in the immediate window
image: /solidworks-api/geometry/transformation/get-coordinate-system-transform/coordinate-system.png
labels: [transform,coordinate system]
---
![Coordinate system in the feature manager tree](coordinate-system.png){ width=450 }

This VBA macro extract the 4x4 [transformation matrix](/solidworks-api/geometry/transformation/) from the selected coordinate system in the feature manager tree.

The comma separated results are output to the immediate (ctrl+G) window of VBA editor.

![Matrix output to the immediate window of VBA editor](maxtrix-output-immediate.png){ width=350 }

{% code-snippet { file-name: Macro.vba } %}
