---
layout: article
title: Select component in feature tree using its name via SOLIDWORKS API
caption: Select Component By Name
description: Example demonstrates how to select component by full name at any level of the assembly in the feature manager tree
image: components-tree.png
labels: [select, component]
---
![Multi-level tree of components](components-tree.png){ width=200 }

This example demonstrates the most performance efficient way to select a component on any level of the assembly by its full name using SOLIDWORKS API.

Name of the component is defined as a path where each level is separated by / symbol. Component instance id is specified with a - symbol (e.g. FirstLevelComp-1/SecondLevelComp-2/TargetComp-1)

Component name can be found in the following dialog:

![Component name in properties dialog](component-name.png){ width=250 }

Refer [Get Component By Name](/solidworks-api/document/assembly/components/get-by-name) example for macro to retrieve the pointer to the component without the selection.

{% code-snippet { file-name: Macro.vba } %}
