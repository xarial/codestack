---
layout: article
title: Get the total length of segments in selected sketch using SOLIDWORKS API
caption: Get Total Length Of Sketch Segments
description: C# example to calculate total length of all non construction geometry sketch segments in the selected sketch using SOLIDWORKS API
image: sketch-total-length.png
labels: [sketch,length]
---
![Total length of the selected sketch segments](sketch-total-length.png){ width=450 }

This C# example of [stand-alone console application](/solidworks-api/getting-started/stand-alone/) to calculate the total length of all segments in the selected sketch using SOLIDWORKS API. Construction geometry sketch segments are excluded from the calculation.

{% code-snippet { file-name: Program.cs } %}
