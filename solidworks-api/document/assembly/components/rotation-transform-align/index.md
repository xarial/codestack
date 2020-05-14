---
layout: article
title: Calculating rotation transformation to align component with direction
caption: Aligning Component With Rotation Transformation
description: VBA example demonstrates hwo to calculate the rotation transformation to align the normal of the component's face with edge direction around the component's origin
image: /solidworks-api/document/assembly/components/rotation-transform-align/rotation-transform.png
labels: [transform,rotation,align]
---
This VBA example demonstrates how to use the [IMathUtility::CreateTransformRotateAxis](https://help.solidworks.com/2017/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IMathUtility~CreateTransformRotateAxis.html) SOLIDWORKS API to rotate the component and align the normal of its face with the direction from the linear edge.

As a precondition select the planar face on the first component in the assembly and linear edge on the second component in the assembly. First component must not be fixed and do not have any mates. As the result first component rotated in a way that its normal is collinear with the direction of the edge. Component is rotated around the origin.

## Explanation

In order to transform the component in the expected way it is required to calculate its transform. For that it is required to find the origin of rotation, rotation vector and an angle.

At first we create vectors of the face normal and edge direction. It is required to apply the transformation of the components to represent vectors in the same coordinate system. The angle between those vectors is a required angle of transformation.

In order to find the vector of rotation it is required to find the vector perpendicular to both normal and direction. This can be achieved by finding the cross product.

Finally point of rotation is an origin of the component transformed to the assembly coordinate system.

{% include img.html src="rotation-transform.png" alt="Rotation transformation parameters" align="center" %}

{% code-snippet { file-name: Macro.vba } %}
