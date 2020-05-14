---
layout: article
title: Determine if selected face is hole or boss using SOLIDWORKS API
caption: Determine If The Selected Face Is Hole Or Boss
description: Example demonstrates how to identify if the selected cylindrical face in SOLIDWORKS part or assembly is internal (i.e. hole) or external (i.e. boss) using SOLIDWORKS API based on the normals of the face.
image: /solidworks-api/geometry/determine-hole-boss/boss-hole.png
labels: [geometry, hole, boss]
---
{% include img.html src="boss-hole.png" width=250 alt="Holes and bosses in the body" align="center" %}

This example demonstrates how to identify if the selected cylindrical face is internal (i.e. hole) or external (i.e. boss) using SOLIDWORKS API.

Select cylindrical face and run the macro. Message box is displayed with the type of the selected face. Macro will work with any face (it is not required for faces to have planar adjacent faces).

### Algorithm

This macro identifies if the face is hole or boss based on the direction of the normal of the face. The normals for the holes are always directed towards the cylinder axis, while the normals for the bosses always directed outwards of the cylinder axis.

Macro finds random point on the face (in this example this is a middle between U and V parameters of the face) and normal at this point. After, the vector between this point and the cylinder origin is calculated. If the angle between this vector and normal is less than 90 degrees (PI / 2) than the normal is directed towards the cylinder axis which means that the face is a hole, otherwise (if angle is greater than 90 degrees (PI / 2)) the face is external (boss).

Please see image below:

{% include img.html src="inner-face-outer-face.png" width=400 alt="Normals for the hole and boss" align="center" %}

{% code-snippet { file-name: Macro.vba } %}