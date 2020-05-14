---
layout: article
title: Modify configuration parameters for components using SOLIDWORKS API
caption: Modify Configuration Parameters For Multiple Components
description: Example demonstrates how to modify parameters of multiple components in the specified configurations (e.g. suppression state) using SOLIDWORKS API
image: /solidworks-api/document/assembly/components/batch-modify-parameters/modify-configurations.png
labels: [parameters, design table, components, configuration]
---
{% include img.html src="modify-configurations.png" width=350 alt="Modify component parameters in configurations" align="center" %}

This example demonstrates how to use parameters (similar to design table parameters) to suppress all components in every configuration except of the active one using SOLIDWORKS API. It is not required to activate configuration or select any components to use the macro.

Multiple components can be modified in a batch mode to improve performance.

{% include_relative Macro.vba.codesnippet %}