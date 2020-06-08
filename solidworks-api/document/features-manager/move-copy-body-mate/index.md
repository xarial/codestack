---
title: Add move-copy body feature with coincident mate using SOLIDWORKS API
caption: Add Move-Copy Body Feature With Mate
description: C# VSTA macro example to create move-copy body feature and add coincident mate between the largest face of the body and front plane using SOLIDWORKS API
image: move-copy-body-mate-pmp.png
labels: [move-copy body,mates]
---
![Move-Copy Body Property Manager Page with mates added](move-copy-body-mate-pmp.png){ width=150 }

C# VSTA macro example which finds the largest planar face of the selected body and inserts move-copy body feature in part and adds coincident mate with Front Plane using SOLIDWORKS API.

* Open part document
* Select any body which contains the planar face
* Run the macro. As the result move-copy body feature is inserted via [IFeatureManager::InsertMoveCopyBody2](http://help.solidworks.com/2016/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeaturemanager~insertmovecopybody2.html) SOLIDWORKS API method. Then coincident mate is added between the largest face of the body and front plane using [IMoveCopyBodyFeatureData::AddMate](http://help.solidworks.com/2016/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IMoveCopyBodyFeatureData~AddMate.html) SOLIDWORKS API method.

{% code-snippet { file-name: SolidWorksMacro.cs } %}
