---
title: Storing parameters in SOLIDWORKS macro feature
caption: Parameters
description: Storing the parameters structure in SOLIDWORKS macro feature using SwEx.MacroFeature framework
toc-group-name: labs-solidworks-swex
order: 1
---
Parameters are any additional metadata required by the macro feature. Currently only primitive types of parameters are supported (i.e. string, bool, double, int etc.)

~~~ cs
public class MacroFeatureParams
{
    public string Parameter1 { get; set; }
    public int Parameter2 { get; set; }
}

//this macro feature has two parameters (Parameter1 and Parameter2)
[ComVisible(true)]
public class MyMacroFeature : MacroFeatureEx<MacroFeatureParams>
{
}
~~~