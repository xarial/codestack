---
layout: article
title: Get the total length of segments in selected sketch using SOLIDWORKS API
caption: Get Total Length Of Sketch Segments
description: C# example to calculate total length of all non construction geometry sketch segments in the selected sketch using SOLIDWORKS API
lang: en
image: /solidworks-api/document/sketch/get-sketch-segments-total-length/sketch-total-length.png
labels: [sketch,length]
---
{% include img.html src="sketch-total-length.png" width=450 alt="Total length of the selected sketch segments" align="center" %}

This C# example of [stand-alone console application](/solidworks-api/getting-started/stand-alone/) to calculate the total length of all segments in the selected sketch using SOLIDWORKS API. Construction geometry sketch segments are excluded from the calculation.

{% include_relative Program.cs.codesnippet %}
