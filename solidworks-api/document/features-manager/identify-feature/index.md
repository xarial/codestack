---
layout: sw-tool
title: Identify SOLIDWORKS API feature definition and specific type
caption: Identify Feature Definition And Specific Type
description: Helper methods allowing to identify the definition and specific type of the selected feature via SOLIDWORKS API and reflection
image: /solidworks-api/document/features-manager/identify-feature/specific-feature-types.png
labels: [reflection, specific feature, feature definition]
categories: sw-tools
group: Developers
---
{% include img.html src="specific-feature-types.png" width=450 alt="Type of specific feature and feature definition of selected feature output to the window" align="center" %}

[IFeature::GetSpecificFeature2](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~GetSpecificFeature2.html) and [IFeature::GetDefinition](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeature~getdefinition.html) SOLIDWORKS API methods return dispatch pointers which in some cases are not easy to identify and cast to specific types.

The following code example allows to output all assignable interfaces for the selected feature's definition and specific feature. The result is output to the *Output* window of VSTA editor.

{% code-snippet { file-name: Macro.cs } %}
