---
layout: sw-tool
title: Macro to scale drawing views based on the geometry size using SOLIDWORKS API
caption: Scale Views Based On Geometry Size
description: VBA macro to scale drawing views in the current sheet based on the geometry size and specified map
lang: en
image: /solidworks-api/document/drawing/views-size-based-scale/drawing-view-scale.png
labels: [scale,size,bounding box]
categories: sw-tools
group: Drawing
---
{% include img.html src="drawing-view-scale.png" width=250 alt="Drawing view scale options" align="center" %}

This VBA macro automatically scales drawing views in the current sheet based on the geometry size and specified matching map.

Map is a collection of instructions which defines the

* Minimum and maximum width of the geometry. Specify * to match any value
* Minimum and maximum height of the geometry. Specify * to match any value
* Scale nominator and denominator if matched

Geometry size is calculated based on the bounding box of visible entities in the drawing view (this includes all the reference geometry, sketch entities, dimensions and other annotations):

{% include img.html src="drawing-view-parameters.png" width=350 alt="Drawing view geometry size parameters" align="center" %}

All drawing views have an offset boundary. This boundary is deducted from the view size in order to get the actual value of the geometry. The value of the boundary is calculated dynamically (2% of the width or height of the sheet, whichever is smaller). This is not a documented value and might change in future by SOLIDWORKS which may affect the calculations in this macro.

{% include img.html src="boundary-offset.png" alt="Boundary offset of drawing view" align="center" %}

Configure the scale map at the beginning of the macro. Specify as many map entries as needed.

~~~ vba
Dim scaleMap As Variant
scaleMap = Array("0-0.1;*;1:1", "0.1-0.2;0.05-0.1;1:2", "another entry", ..., "last entry")
~~~

Each entry must follow the predefined format:

~~~
"[minWidth]-[maxWidth];[minHeight]-[maxHeight];[scaleNom]:[scaleDenom]"
~~~

* All values for width and height are in meters
* Specify * to allow any width or height

In the example below

~~~ vba
Array("0-0.1;*;1:1", "0.1-0.2;0.05-0.1;1:2")
~~~

* All drawing views with width up to 100 mm and any height will be set to 1:1 scale
* All drawing views with width between 100 mm to 200 mm and height between 50 mm to 100 mm will be set to 1:2 scale

{% include_relative Macro.vba.codesnippet %}