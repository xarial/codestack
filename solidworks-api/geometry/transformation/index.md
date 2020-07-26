---
title: Using transformations in SOLIDWORKS API
caption: Transformation
description: Applying and reading transformation (components, bodies, sketches etc.) using SOLIDWORKS API
order: 2
labels: [transform,math]
---
Transformation in SOLIDWORKS API is represented in the [IMathTransform](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imathtransform.html) interface. This is a 4x4 transformation matrix holding the information about transform, scale and translation.

Matrix array is represented with 16 doubles (0 to 15) which are aligned in the following order

|||||
|-|-|-|-|
|0|1|2|13|
|3|4|5|14|
|6|7|8|15|
|9|10|11|12|

0-8 - rotational part of the matrix

9-11 - translation part of matrix (x, y, z)

12 - scale factor

13-15 - not used

Matrix describe the orientation and translation of various elements in SOLIDWORKS, such as

* Components positions in the assembly
* Bodies relative movements
* Relation between sketch coordinate system and model coordinate system
* Camera orientation and model view rotation

Identity matrix which represents no rotation, scale or transform will be equal to

**Visual Basic**

~~~ vb
Dim dMatrix(15) As Double
dMatrix(0) = 1: dMatrix(1) = 0: dMatrix(2) = 0: dMatrix(3) = 0
dMatrix(4) = 1: dMatrix(5) = 0: dMatrix(6) = 0: dMatrix(7) = 0
dMatrix(8) = 1: dMatrix(9) = 0: dMatrix(10) = 0: dMatrix(11) = 0
dMatrix(12) = 1: dMatrix(13) = 0: dMatrix(14) = 0: dMatrix(15) = 0
~~~

**C#**

~~~ cs
var matrix = new double[]
{
    1, 0, 0, 0,
    1, 0, 0, 0,
    1, 0, 0, 0,
    1, 0, 0, 0
};
~~~

[IMathUtility](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imathutility.html) is a SOLIDWORKS API utility interface providing the access to operations related to composing the transformation based on input parameters (such as rotation angles, translation, raw data).

The following interfaces are usually used while calculation of transformations and translations:

[IMathVector](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imathvector.html)
[IMathPoint](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imathpoint.html)

This example contains articles and tutorials explaining the use of transformation matrices.