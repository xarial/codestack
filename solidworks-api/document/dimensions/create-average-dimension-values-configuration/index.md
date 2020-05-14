---
layout: sw-tool
title: SOLIDWORKS macro to create configuration with average dimension values
caption: Create Configuration With Average Dimension Values
description: Macro will create child configuration where all the dimension will be set to average value based on the minimum and maximum values of the tolerance
image: /solidworks-api/document/dimensions/create-average-dimension-values-configuration/sw-dimension-tolerance.png
labels: [average, configuration, dimension, solidworks api, tolerance, utility]
categories: sw-tools
group: Model
redirect-from:
  - /2018/03/solidworks-api-dimensions-average-dims.html
---
This macro will create child configuration where all the dimension will be set to average value based on the minimum and maximum values of the tolerance using SOLIDWORKS API.

{% include img.html src="sw-dimension-tolerance.png" width=400 alt="Dimension Tolerance/Precision group in property manager page" align="center" %}

{% code-snippet { file-name: Macro.vba } %}
