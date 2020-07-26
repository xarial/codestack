---
title: SOLIDWORKS macro to change configuration specific properties for component in pattern
caption: Change Configuration Specific Properties For Component In Pattern
description: Example demonstrates how to change the configuration specific properties (use same configuration as pattern seed component or use named configuration) of the component in the pattern using SOLIDWORKS API
image: component-config-specific-properties.png
labels: [assembly, spattern, configuration, seed]
---
![Configuration specific properties for the seed component of the sketch driven pattern](component-config-specific-properties.png)

This macro example demonstrates how to change the following configuration specific properties using SOLIDWORKS API.

* Use same configuration as pattern seed component
* Use named configuration

[IAssemblyDoc::CompConfigProperties5](https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iassemblydoc~compconfigproperties5.html) SOLIDWORKS API is used to modify the multiple properties at a time for the selected component.

In the instance component of the pattern (e.g. Sketch Driven Pattern)

{% code-snippet { file-name: Macro.vba } %}
