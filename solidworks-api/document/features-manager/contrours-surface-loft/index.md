---
layout: article
title: SOLIDWORKS API to create surface loft feature via contours
caption: Create Surface Loft Feature Via Contours
description: Example demonstrates how to create surface loft feature from the contours as the profiles using SOLIDWORKS API
image: /solidworks-api/document/features-manager/contrours-surface-loft/lofted-surface-sketch-contours.png
labels: [surface, loft, contour]
---
![Lofted surface feature using sketch contours as the profiles](lofted-surface-sketch-contours.png){ width=500 }

This example demonstrates how to create surface loft feature from the contours as the profiles using SOLIDWORKS API.

Sketch segments are not accepted entities for the profiles in the surface loft feature. This means if only several segments from the sketch need to be used for profiles (instead of the entire sketch) it is not possible to create a feature by selecting the sketch segments. It is required to use sketch contours instead.

Sketch segments are not supported from the User Interface as well. When segment is selected the following selection manager is displayed allowing to select the open or closed loop.

![Selection manager while selecting the profile](selection-manager.png){ width=250 }

* Open part and select sketch segments for profile. Any types of sketch segments are supported (spline, line, arc etc.). There might be multiple sketch segments in the sketch and only several can be selected for the profile. Segments can be in different sketches as well.
* Macro will find the corresponding sketch contour for each sketch segment
* Macro will create surface loft feature with the corresponding sketch contours

> This macro is not an optimal performance code for finding sketch contours of segments within the same sketch as it will do a full traversal of all sketch segments within the sketch to find the corresponding contour for individual sketch segments. Modify the macro to find multiple sketch contours at a time within one traversal loop avoiding repetition.

{% code-snippet { file-name: Macro.vba } %}
