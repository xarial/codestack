---
title: Managing of Edit Bodies in SOLIDWORKS macro feature
caption: Edit Bodies
description: Managing of Edit Bodies in SOLIDWORKS macro feature using SwEx.MacroFeature framework
toc-group-name: labs-solidworks-swex
order: 3
---
Edit bodies are input bodies which macro feature will acquire. For example when boss-extrude feature is created using the merge bodies option the solid body it is based on became a body of the new boss-extrude. This could be validated by selecting the feature in the tree which will select the body as well. In this case the original body was passed as an edit body to the boss-extrude feature.

~~~ cs
public class MacroFeatureParams
{
    [ParameterEditBody]
    public IBody2 InputBody { get; set; }
}
~~~

If multiple input bodies are required it could be either specified in different properties

~~~ cs
public class MacroFeatureParams
{
    [ParameterEditBody]
    public IBody2 EditBody1 { get; set; }

    [ParameterEditBody]
    public IBody2 EditBody2 { get; set; }
}
~~~

or as list

~~~ cs
public class MacroFeatureParams
{
    [ParameterEditBody]
    public List<IBody2> EditBodies { get; set; }
}
~~~