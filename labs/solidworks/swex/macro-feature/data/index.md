---
layout: article
title: Storing data (parameters, bodies, selection) in SOLIDWORKS macro feature
caption: Data
description: Storing the parameters, metadata, dimensions, selections in the SOLIDWORKS macro feature using SwEx.MacroFeature framework
toc-group-name: labs-solidworks-swex
order: 3
---
Macro feature can store additional metadata and entities. The data includes

* Parameters
* Selections
* Edit bodies
* Dimensions

Required data can be defined within the macro feature data model. Special parameters (such as selections, edit bodies or dimensions) should be decorated with appropriate [attributes](https://docs.codestack.net/swex/macro-feature/html/N_CodeStack_SwEx_MacroFeature_Attributes.htm), all other properties will be considered as parameters.

Data model is used both as input and output of macro feature. Parameters can be accessed via [GetParameters](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_1_GetParameters.htm) method and also passed to [OnRebuild](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_1_OnRebuild.htm) handler. Parameters can be modified by calling the [SetParameters](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_1_SetParameters.htm) method.

~~~ cs
public class MacroFeatureParams
{
    // text metadata
    public string TextParameter { get; set; }
    
    // boolean metadata
    public bool ToggleParameter { get; set; }

    // any dependency selection
    [ParameterSelection]
    public IFace2 FaceSelectionParameter { get; set; }

    // edit body - base body which macro feature is modifying
    [ParameterEditBody]
    public IBody2 InputBody { get; set; }

    // macro feature dimension. Value of the dimension will be sync with the proeprty
    [ParameterDimension(swDimensionType_e.swLinearDimension)]
    public double LinearDimension { get; set; }
}

[ComVisible(true)]
public class MyMacroFeature : MacroFeatureEx<MacroFeatureParams>
{
}
~~~