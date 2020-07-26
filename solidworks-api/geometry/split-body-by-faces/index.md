---
layout: sw-tool
title: SOLIDWORKS Macro to Split Body By Faces using SOLIDWORKS API
caption: Split Body By Faces
description: Macro splits the selected surface or solid body by faces creating individual sheet body for each face using SOLIDWORKS API
image: feature-manager-tree-split-faces.png
labels: [split,body,faces]
group: Geometry
---
![Feature Manager Tree with sheet bodies for each face](feature-manager-tree-split-faces.png){ width=250 }

This macro creates individual surface (sheet) body for each face of the selected solid or surface body using the [IModeler::CreateSheetFromFaces](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeler~createsheetfromfaces.html) SOLIDWORKS API method.

{% code-snippet { file-name: Macro.vba } %}

For more advanced functionality (supporting parametric approach) refer the [Geomtery++ Split Body By Faces feature](/labs/solidworks/geometry-plus-plus/user-guide/split-body-by-faces/)
