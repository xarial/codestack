---
layout: article
title: Select component in feature tree using its name via SOLIDWORKS API
caption: Select Component By Name
description: Example demonstrates how to select component by full name at any level of the assembly in the feature manager tree
lang: en
image: /solidworks-api/document/selection/select-component-by-name/components-tree.png
labels: [select, component]
---
{% include img.html src="components-tree.png" width=200 alt="Multi-level tree of components" align="center" %}

This example demonstrates the most performance efficient way to select a component on any level of the assembly by its full name using SOLIDWORKS API.

Name of the component is defined as a path where each level is separated by / symbol. Component instance id is specified with a - symbol (e.g. FirstLevelComp-1/SecondLevelComp-2/TargetComp-1)

Component name can be found in the following dialog:

{% include img.html src="component-name.png" width=250 alt="Component name in properties dialog" align="center" %}

Refer [Get Component By Name]({{ "/solidworks-api/document/assembly/components/get-by-name" | relative_url }}) example for macro to retrieve the pointer to the component without the selection.

{% include_relative Macro.vba.codesnippet %}
