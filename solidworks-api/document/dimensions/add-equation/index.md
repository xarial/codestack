---
layout: article
title: Add equation to dimension using SOLIDWORKS API
caption: Add Equation To Dimension
description: Example will modify the value of the selected dimension and sets its value to be equal to the equation
image: sw-dimension-equation.png
labels: [dimension, solidworks api, equation, example]
redirect-from:
  - /2018/03/solidworks-api-dimensions-add-equation-to-dim.html
---
This example will modify the value of the selected dimension and sets its value to be equal to the equation using SOLIDWORKS API:

> sin(0.5) * 2 + (10 - 5)

![Equation in dimension](sw-dimension-equation.png){ width=320 height=200 }

[IEquationMgr](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IEquationMgr.html) SOLIDWORKS API interface should be used to manage equations in SOLIDWORKS document.

{% code-snippet { file-name: Macro.vba } %}
