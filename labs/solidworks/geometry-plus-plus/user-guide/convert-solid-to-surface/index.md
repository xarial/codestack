---
layout: article
title: Convert Solid To Surface feature in Geometry++
caption: Convert Solid To Surface
description: Converts solid bodies to surface bodies in SOLIDWORKS part document preserving the parametric functionality
image: /labs/solidworks/geometry-plus-plus/user-guide/convert-solid-to-surface/icon.png
toc_group_name: labs-solidworks-geometry-plus-plus
---
This feature converts solid bodies to surface bodies.

{% include img.html src="convert-solid-to-surface-page.png" alt="Convert solid body to surface body property manager page" align="center" %}

* Select solid body or bodies to convert
* Click green tick

New feature is added to the feature manager tree.

{% include img.html src="solid-to-surface-feature.png" alt="Solid to surface feature in the feature manager tree" align="center" %}

All selected solid bodies are replaced by corresponding surface bodies. Feature is fully parametric. If some of the base features which were forming solid bodies modified the surface body feature is modified as well.
