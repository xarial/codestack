---
layout: article
title: Convert meters to fraction inches using SOLIDWORKS API
caption: Convert Meters To Fraction Inches
description: VBA macro to convert value in system units (meters) to the fraction inches using SOLIDWORKS API
image: /solidworks-api/application/convert-to-fraction-inches/dimension-fractions.png
labels: [inches,fraction,conversion,unit]
---
This VBA macro converts the value specified in system units (meters) into the fraction inches with specified denominator.

For example the value of 0.112713 with denominator 16 would be converted to 4 7/16"

Configure the parameters as per the settings below:

~~~ vb
Const DENOMINATOR As Integer = 16 'denominator value
Const ROUND_TO_NEAREST_FRACTION As Boolean = True 'True to round to nearest fraction, False to not
~~~

The result and available options are equal to the dimension property manager page in SOLIDWORKS

{% include img.html src="dimension-fractions.png" alt="Override units option in the Dimension proeprty manager page" align="center" %}

{% include_relative Macro.vba.codesnippet %}
