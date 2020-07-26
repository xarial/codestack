---
title: Fix or float component in active or all configurations using SOLIDWORKS API
caption: Fix/Float In This Or All Configurations
description: Example demonstrates a workaround for missing SOLIDWORKS API for fixing or floating the component in the active or all configuration
image: component-fix-options.png
labels: [fix, float, component, workaround]
---
![Options to fix component](component-fix-options.png)

This VBA example demonstrates a simple workaround for missing SOLIDWORKS API to fix or float the component in active configuration only. [IAssemblyDoc::FixComponent](https://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iassemblydoc~fixcomponent.html) only fixes the components in all configurations.

Create an example assembly with 2 configurations and 4 instances of the component, where first 2 instances are floating in both configurations, while last 2 instances are fixed in both configuration.

![Initial state of example](component-initial-state.png)

As the result of running this macro components will be changed to the following result:

![Result of running the macro](component-fix-result.png)

{% code-snippet { file-name: Macro.vba } %}
