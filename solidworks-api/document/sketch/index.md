---
layout: article
title: Working with sketch using SOLIDWORKS API
caption: Sketch
description: Working with 2D and 3D sketches (adding and reading segments, transformations, feature creation) using SOLIDWORKS API
order: 6
labels: [sketch,draw]
---
Sketch is a 3D or 3D layout in SOLIDWORKS parts, assemblies and drawing. In most cases sketch is used as a profile for generating 3D elements (extrudes, cuts, lofts etc.).

Sketch is a feature and it is managed via [ISketch](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketch.html) interface in SOLIDWORKS API.

Sketch can contain sketch segments and sketch points as well as annotations (dimensions, notes, etc.).

2D sketch uses 2D coordinate system (X, Y) to position its elements. This coordinate system not always matches the global coordinate system. Which means that the coordinates of elements found in the sketch are relative to 2D coordinate system and need to be transformed to model space if required.

This section contains various macros and code examples of working with sketches, adding and removing segments and points, creating new sketches, calculating transformation using SOLIDWORKS API.