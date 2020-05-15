---
layout: article
title: Create temp solid body box using SOLIDWORKS API and IModeler interface
caption: Create Box Body
description: VBA example to create a temp body by center point, direction and size using SOLIDWORKS API and IModeler interface
image: box-body.png
labels: [primitive,box,temp body,modeler]
---
![Box body](box-body.png){ width=250 }

This VBA example demonstrates how to create and display temp body by providing the coordinate of center of base face, direction, width, length and height using SOLIDWORKS API.

Macro stops the execution and displays the body. Continue macro execution to destroy the temp body.

{% code-snippet { file-name: Macro.vba } %}
