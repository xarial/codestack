---
layout: sw-tool
title: Macro create precise part bounding box using SOLIDWORKS API
caption: Create Precise Bounding Box
description: Macro creates a precise bounding box in the part document using SOLIDWORKS API
image: /solidworks-api/geometry/precise-bounding-box/precise-bounding-box.png
labels: [bonding box, extreme points]
group: Part
---
![Precise bounding box in the part document](precise-bounding-box.png){ width=250 }

As per *Remarks* section of [IPartDoc::GetPartBox](http://help.solidworks.com/2016/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ipartdoc~getpartbox.html) method (or other BoundingBox APIs) in SOLIDWORKS API Help Documentation

> The values returned are approximate and should not be used for comparison or calculation purposes. Furthermore, the bounding box may vary after rebuilding the model

To calculate the precise bounding box it is required to find the extreme points of each body in XYZ directions via [IBody2::GetExtremePoint](http://help.solidworks.com/2016/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibody2~getextremepoint.html)

The following macros will calculate the bounding box, width, height and length of the active part document using both approaches of SOLIDWORKS API.

As the result 3D Sketch with bounding box is created.

### Precision

Bounding boxes calculated approximately might be more than 10% inaccurate. For the following [example part](bbox-precision.SLDPRT) the difference between the bounding boxes volumes equal to 14%. The following images show the differences (green box is a precise calculation and red box is an approximate calculation):

![Front View](bbox-front-view.png){ width=250 }

![Top View](bbox-top-view.png){ width=250 }

![Right View](bbox-right-view.png){ width=250 }

> The precise bounding box calculated by extreme points is exactly equal to the bounding box created by [bounding box feature](http://help.solidworks.com/2018/English/WhatsNew/t_bounding_box_for_part_assem.htm) added in SOLIDWORKS 2018

### Performance

Extraction of approximate box is more than 300 times quicker. For a single body part approximate calculation of bounding box took 0.016ms, while it took 5.57 ms for precise calculation of the same part. For multi-body part of 63 bodies it took 0.018ms for approximate calculations and 16.68 ms for precise calculations.

As a summary on avarage it would be possible to calculate more than 60000 approximate bounding boxes per second and only about 50 precise bounding boxes per second (more than 1000 times difference)

### Calculating precise bounding box via extreme points

{% code-snippet { file-name: GetPreciseBoundingBox.vba } %}

### Calculating approximate bounding box

{% code-snippet { file-name: GetBoundingBox.vba } %}
