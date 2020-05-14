---
layout: article
title: Understanding transforms in sketches while using SOLIDWORKS API
caption: Understanding Sketch Transforms
description: Explanation of model to sketch and sketch to model transformations in SOLIDWORKS API to properly calculate the coordinates of sketch segments
image: /solidworks-api/document/sketch/transform/sketch-coordinate-systems.png
labels: [transform,sketch]
---
When working with sketch segments (e.g. line, arc, etc.) or points it is important to consider the fact that the coordinates values returned from SOLIDWORKS API such as [ISketchPoint::X](http://help.solidworks.com/2017/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchPoint~X.html) property are relative to the local sketch coordinate system.

Those values will match for 3D Sketches or 2D sketches created on Front plane (if not moved), but will be different in other cases.

As shown on the following picture the value of the point is displayed as { -50, 10, 0 } for the local sketch coordinate system (in the sketch point property manager page) and as the { -50, 0, -10 } for the global coordinate system (in the SOLIDWORKS status bar). This difference is caused by the fact that 2D sketch is created on the Top Plane.

{% include img.html src="global-local-coordinates.png" width=450 alt="Different values for the local and global coordinate systems." align="center" %}

Local coordinate system of 2D sketch is displayed with red X and Y arrows when activating the sketch. And global coordinate system is represented with red, green and blue triad in the bottom right corner of SOLIDWORKS model window.

{% include img.html src="sketch-coordinate-systems.png" width=350 alt="Local sketch coordinate system and global coordinate system" align="center" %}

## Reading the local coordinates from sketch point

The following macro reads the selected sketch point coordinate relative to the local sketch coordinate system and outputs it to the immediate Window of SOLIDWORKS.

{% include img.html src="coordinate-output.png" width=350 alt="Extracted coordinate of sketch point" align="center" %}

* Create a sketch on the Front Plane and create a sketch point
* Select this point
* Run the macro and compare with the global coordinate value (result is printed in meters)
* Values will match

{% include img.html src="sketch-point-coordinate.png" width=350 alt="Sketch point global coordinate" align="center" %}

* Create new sketch on any plane but Front Plane (e.g. Top Plane)
* Repeat the steps above
* Now coordinates do not match.

{% code-snippet { file-name: NoTransformsMacro.vba } %}

## Retrieving the global coordinates from sketch point

In order to find the value of the coordinate relative to the global coordinate system it is required to find the sketch to model [transformation matrix]({{ "/solidworks-api/geometry/transformation/" | relative_url }}) via [ISketch::ModelToSketchTransform](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketch~ModelToSketchTransform.html) SOLIDWORKS API property and apply this to the point coordinate.

Below macro can be used to perform the steps from the previous paragraph, but now the extracted coordinates will match the values in the global coordinate system.

{% code-snippet { file-name: TransformSketchCoordinateMacro.vba } %}

## Creating point in sketch from global coordinates

Inversed transformation should be used when it is required to create a sketch point in the 2D sketch based on the global coordinate value. The following example inserts a sketch point into an active sketch based on a XYZ value.

{% code-snippet { file-name: CreateSketchPointFromGlobalCoordinate.vba } %}
