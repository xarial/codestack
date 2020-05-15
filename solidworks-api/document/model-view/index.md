---
layout: article
title: Manipulating model views using SOLIDWORKS API
caption: Model Views
description: Collection of articles and code examples for working with 3D model views using SOLIDWORKS API
order: 3
---
Model view is a 3D snapshot of SOLIDWORKS model visible to the user. SOLIDWORKS API provides the [IModelView](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelView.html) interface to perform manipulation and data extraction from views.

Model views can be transformed (scaled, rotated, moved) to change the orientation of the model.

Multiple view can be presented in the document to represent various states of the model. For example the motion study tab create new views to render the motion specific user interface elements.

This section contains examples and macros for manipulating the model views using SOLIDWORKS API.