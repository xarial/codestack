---
layout: article
title: Crop Bodies feature in Geometry++
caption: Crop Bodies
description: Feature allows trimming surface or solid or multiple surfaces or solids using sketch or sketch contour in SOLIDWORKS part document
image: /labs/solidworks/geometry-plus-plus/user-guide/crop-bodies/icon.png
toc_group_name: labs-solidworks-geometry-plus-plus
redirect_from:
  - /labs/solidworks/geometry-plus-plus/user-guide/region-trim-surface/
---
This command allows trimming surface or solid (target bodies) using sketches or sketch regions (trimming tools).

Multiple target bodies and trimming tools are supported.

* Select surface or solid body or bodies from the graphics area or from the feature tree. Box selection is supported
* Select sketches or sketch regions (requires setting of solid works filter) to trim the surface. Tool will keep the surface geometry which resides within the sketch region.
Feature will trim the surface perpendicular to the corresponding trim tool sketches normals.

{% include img.html src="crop-bodies-page.png" width=500 alt="Crop bodies property manager page" align="center" %}

Once selection completed and green tick is clicked new feature is added to the feature manager tree.

{% include img.html src="cropped-bodies.png" width=500 alt="Original bodies and resulted cropped geometry" align="center" %}

Original bodies are acquired by new feature. The bodies outside of the region will be removed by macro feature.

{% include img.html src="crop-body-feature.png" height=450 alt="Crop bodies feature in the feature manager tree" align="center" %}

Feature can be edited, removed and rollbacked as any other feature.
