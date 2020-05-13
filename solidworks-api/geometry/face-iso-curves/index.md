---
layout: article
title: Macro to create ISO curves for face using SOLIDWORKS API
caption: Create ISO Curves For Face
description: Example demonstrates how to find specified number of iso-curves in the u and v bounds of the selected face using SOLIDWORKS API
lang: en
image: /solidworks-api/geometry/face-iso-curves/iso-curves-wire-body.png
labels: [curve, evaluate, geometry, macro, iso, uv, trimmed curve, vba]
---
{% include img.html src="iso-curves-wire-body.png" height=300 alt="Preview of iso curves of the face" align="center" %}

This example demonstrates how to find specified number of iso-curves in the u and v bounds of the selected face using SOLIDWORKS API.

* Select the face and run the macro
* Iso curves are previewed and macro execution stops
* Continue the macro to clear the preview

Number of iso curves in u and v direction can be changed in the following snippet

~~~ vb
Dim vCurves As Variant
vCurves = GetIsoCurves(swFace, <Number of curves in u direction>, <Number of curves in v direction>)
~~~

Optionally macro allows to create curves in the 3D Sketch.

{% include img.html src="iso-curves-sketch.png" height=300 alt="Sketch created for iso curves of the face" align="center" %}

This option can be enabled by setting *CREATE_SKETCH* constant to *True* at the beginning of the macro:

~~~ vb
Const CREATE_SKETCH As Boolean = True
~~~

**Macro:**

{% include_relative Macro.vba.codesnippet %}
