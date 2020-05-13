---
layout: article
title: Fix or float component in active or all configurations using SOLIDWORKS API
caption: Fix/Float In This Or All Configurations
description: Example demonstrates a workaround for missing SOLIDWORKS API for fixing or floating the component in the active or all configuration
lang: en
image: /solidworks-api/document/assembly/components/fix-float/component-fix-options.png
labels: [fix, float, component, workaround]
---
{% include img.html src="component-fix-options.png" alt="Options to fix component" align="center" %}

This VBA example demonstrates a simple workaround for missing SOLIDWORKS API to fix or float the component in active configuration only. [IAssemblyDoc::FixComponent](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iassemblydoc~fixcomponent.html) only fixes the components in all configurations.

Create an example assembly with 2 configurations and 4 instances of the component, where first 2 instances are floating in both configurations, while last 2 instances are fixed in both configuration.

{% include img.html src="component-initial-state.png" alt="Initial state of example" align="center" %}

As the result of running this macro components will be changed to the following result:

{% include img.html src="component-fix-result.png" alt="Result of running the macro" align="center" %}

{% include_relative Macro.vba.codesnippet %}
