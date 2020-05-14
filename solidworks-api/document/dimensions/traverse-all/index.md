---
layout: article
title: Traverse all dimensions of component or model using SOLIDWORKS API
caption: Traverse All Dimensions
description: VBA macro which traverses all dimensions of all features in the selected component or active document using SOLIDWORKS API and outputs the dimension name and value to the output Window
image: /solidworks-api/document/dimensions/traverse-all/dimensions.png
labels: [dimension,display dimension,traverse]
---
{% include img.html src="dimensions.png" alt="Dimensions in the sketch of weldment feature" align="center" %}

This VBA macro demonstrates how to traverse all dimensions of the features from active SOLIDWORKS document or component (if selected) in the assembly using SOLIDWORKS API.

Macro will output the name of the dimension and the value in the current system units into the Immediate Window of VBA.

~~~
D1@Sketch1=0.15
D2@Sketch1=2.0
RI@Sketch11=0.008
~~~

> The macro will exclude all duplicate dimensions as in some cases (e.g. weldment features) the same dimension may be present in the sketch and in the structural member feature as the same time

{% include_relative Macro.vba.codesnippet %}
