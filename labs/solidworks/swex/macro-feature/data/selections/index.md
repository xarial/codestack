---
layout: article
title: Managing selection of SOLIDWORKS macro feature
caption: Selections
description: Managing selections of SOLIDWORKS macro feature using the SwEx.MacroFeature framework
toc_group_name: labs-solidworks-swex
order: 2
---
~~~ cs
public class MacroFeatureParams
{
    //selection parameter of any entity (e.g. face, edge, feature etc.)
    [ParameterSelection]
    public object AnyEntity { get; set; }

    //selection parameter of body
    [ParameterSelection]
    public IBody2 Body { get; set; }

    //selection parameter of array of faces
    [ParameterSelection]
    public List<IFace2> Faces { get; set; }
~~~

Parameter properties can be specified either using the direct SOLIDWORKS type or as object if type is unknown. List of selections is also supported.

[OnRebuild](https://docs.codestack.net/swex/macro-feature/html/M_CodeStack_SwEx_MacroFeature_MacroFeatureEx_OnRebuild.htm) handler will be called if any of the selections have changed.