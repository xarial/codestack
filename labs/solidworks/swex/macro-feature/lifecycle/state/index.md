---
layout: article
title: Handling the SOLIDWORKS macro feature state update in SwEx.MacroFeature framework
caption: State
description: Updating state of the macro feature on the environment change (selection, rebuild, suppress etc.) 
lang: en
toc_group_name: labs-solidworks-swex
order: 3
---
This handler is called every time state of the feature is changed. It should be used to provide additional security for macro feature.

~~~ cs
protected override swMacroFeatureSecurityOptions_e OnUpdateState(ISldWorks app, IModelDoc2 model, IFeature feature)
{
    //disallow editing or suppressing of the feature
    return swMacroFeatureSecurityOptions_e.swMacroFeatureSecurityCannotBeDeleted 
                | swMacroFeatureSecurityOptions_e.swMacroFeatureSecurityCannotBeSuppressed;
}
~~~