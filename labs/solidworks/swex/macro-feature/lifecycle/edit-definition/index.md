---
layout: article
title: Edit SOLIDWORKS macro feature definition
caption: Edit Definition
description: Edit definition of SOLIDWORKS macro feature using SwEx.MacroFeature framework
toc-group-name: labs-solidworks-swex
order: 2
---
Edit definition allows to modify the parameters of an existing feature. Edit definition is called when *Edit Feature* command is clicked form the feature manager tree.

![Edit Feature Command](menu-edit-feature.png){ width=250 }

The typical workflow which should be followed when feature is edited

* Get the definition of the feature via [IFeature::GetDefinition](http://help.solidworks.com/2016/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeature~getdefinition.html)
* Rollback the feature in the tree via [IMacroFeatureData::AccessSelections](http://help.solidworks.com/2016/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IMacroFeatureData~AccessSelections.html). This will ensure that all the feature selections and edit bodies are available.
* Get the parameters of current macro feature via [GetParameters](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_1_GetParameters.htm)
* Create user interface and allow user to edit parameters. The recommended way to use Property Manager Pages to have a native look and feel of your feature. Use [SwEx.PMPage](/labs/solidworks/swex/pmpage/) framework for simplified way of creating property manager pages.
* Once user interface is closed
    * If OK is clicked, than set modified parameters via [SetParameters](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_1_SetParameters.htm) method and apply the changes to the macro feature via [IFeature::ModifyDefinition](http://help.solidworks.com/2016/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IFeature~ModifyDefinition.html) this step will also rollforward the macro feature in the tree.
    If *Cancel* is clicked undo the modifications via [IMacroFeatureData::ReleaseSelectionAccess](http://help.solidworks.com/2016/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IMacroFeatureData~ReleaseSelectionAccess.html)

{% code-snippet { file-name: EditMacroFeatureDefinition.cs } %}

It is important to use the same pointer to [IMacroFeatureData](http://help.solidworks.com/2016/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imacrofeaturedata.html) while calling the [IMacroFeatureData::AccessSelections](http://help.solidworks.com/2016/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IMacroFeatureData~AccessSelections.html), [GetParameters](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_1_GetParameters.htm), [SetParameters](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_1_SetParameters.htm), [IFeature::ModifyDefinition](http://help.solidworks.com/2016/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IFeature~ModifyDefinition.html) and [IMacroFeatureData::ReleaseSelectionAccess](http://help.solidworks.com/2016/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IMacroFeatureData~ReleaseSelectionAccess.html) methods.