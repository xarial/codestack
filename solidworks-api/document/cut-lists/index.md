---
layout: article
title: Managing cut-list bodies using SOLIDWORKS API
caption: Cut-Lists
description: Automating cut-list bodies (weldment and sheet metal) using SOLIDWORKS API
lang: en
order: 11
image: /images/codestack-snippet.png
labels: [cut-list,weldment,sheet metal]
---
Cut-list bodies got generated from the sheet metal and weldment bodies in SOLIDWORKS. Although those bodies are still managed via [IBody2](http://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibody2.html) SOLIDWORKS API interface they enable additional functionality compared to regular bodies:

* Cut-list bodies are grouped in the cut-list folders by geometry
* Cut-list folders (group of bodies) can have custom properties and auto generated properties (such as length, thickness etc.)

Custom properties could be automated by calling the [IFeature::CustomPropertyManager](http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~CustomPropertyManager.html) property for the cut-list folder item.

Cut-lists are one of the most common elements of SOLIDWORKS API automation. Explore the examples of this section for macros and code snippets for accessing cut-lists data programmatically.
