---
layout: article
title: Get type of cylindrical face using SOLIDWORKS API
caption: Get Type Of Cylindrical Face
description: Macro identifies the type of the selected simple cylindrical face (through all hole, blind hole or external hole) using SOLIDWORKS API based on the loops type
image: /solidworks-api/geometry/cylindrical-face-type/cylindrical-faces-types.png
labels: [geometry, face, hole, outer, inner]
---
{% include img.html src="cylindrical-faces-types.png" height=250 alt="Types of cylindrical faces" align="center" %}

This macro identifies the type of the selected simple cylindrical face (through all hole, blind hole or external hole) based on the loops type using SOLIDWORKS API.

Macro will only work with cylindrical faces whose adjacent faces are planar faces and upper and lower boundaries of the cylinder are closed circular edges.

### Algorithm

Macro traverses the loops of coedges of upper and lower boundary edges. If there is at least one internal loop that means that selected face is a hole, otherwise it is an external boss. If both of the boundary loops are internal that means that the hole is through all, if one boundary loop is external but other is internal that means that the selected face is a blind hole (i.e. not a through all hole).

{% code-snippet { file-name: Macro.vba } %}
