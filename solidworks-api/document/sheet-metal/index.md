---
layout: article
title: Automation Sheet Metal parts using SOLIDWORKS API
caption: Sheet Metal
description: Using SOLIDWORKS API to manipulate sheet metal features
order: 10
image: /images/codestack-snippet.png
labels: [sheet metal,bend,fold]
---
SOLIDWORKS API provide number of methods and interface for manipulating sheet metal features in part documents: [IBaseFlangeFeatureData](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IBaseFlangeFeatureData.html), [IBendsFeatureData](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibendsfeaturedata_members.html), [ISketchedBendFeatureData](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isketchedbendfeaturedata.html) etc.

All the specific feature data could be retrieved via calling the [IFeature::GetDefinition](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~GetDefinition.html) SOLIDWORKS API on the corresponding sheet metal feature.

Explore this section to find useful macros and code examples for automation and enhancement of sheet metal functionality in SOLIDWORKS.
