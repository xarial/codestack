---
layout: article
title: Get the pointer to component from name using SOLIDWORKS API
caption: Get Component By Name
description: Example demonstrates how to get the pointer to the component at any level of the assembly from its full name
image: /solidworks-api/document/assembly/components/get-by-name/components-tree.png
labels: [select, component]
---
{% include img.html src="components-tree.png" width=200 alt="Multi-level tree of components" align="center" %}

This example demonstrates how to retrieve the pointer to the [IComponent2](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.icomponent2.html) SOLIDWORKS API method on any level of the assembly from its full name hierarchy.

Name of the component is defined as a path where each level is separated by / symbol. Component instance id is specified with a - symbol (e.g. FirstLevelComp-1/SecondLevelComp-2/TargetComp-1)

Component name can be found in the following dialog in SOLIDWORKS User Interface:

{% include img.html src="component-name.png" width=250 alt="Component name in properties dialog" align="center" %}

Refer [Select Component By Name]({{ "/solidworks-api/document/selection/select-component-by-name" | relative_url }}) example for an alternative way of selecting the component by name.

{% include_relative Macro.vba.codesnippet %}
