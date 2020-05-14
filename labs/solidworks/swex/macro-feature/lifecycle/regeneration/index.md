---
layout: article
title: Handling Regeneration method of SOLIDWORKS macro feature
caption: Regeneration
description: Handling regeneration event of SOLIDWORKS macro feature and returning bodies or errors to drive the behavior using SwEx.MacroFeature framework
toc_group_name: labs-solidworks-swex
order: 1
---
This handler called when feature is being rebuilt (either when regenerate is invoked or when the parent elements have been changed).

Use [MacroFeatureRebuildResult](https://docs.codestack.net/swex/macro-feature/html/T_CodeStack_SwEx_MacroFeature_Base_MacroFeatureRebuildResult.htm) class to generate the required output.

Feature can generate the following output

{% code-snippet { file-name: RegenerationResults.cs } %}

Use [IModeler](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeler.html) interface if feature needs to create new bodies. Only temp bodies can be returned from the regeneration method.

Use extension methods available in the [IModelerExtension](https://docs.codestack.net/swex/macro-feature/html/T_SolidWorks_Interop_sldworks_ModelerEx.htm) class.