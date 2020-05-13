---
layout: article
title: VBA Macro calls Show All Components command from SOLIDWORKS API
caption: Show All Components (Show With Dependents)
description: Example demonstrates how to call the Show With Dependents command for components or assembly using SOLIDWORKS API
lang: en
image: /solidworks-api/document/assembly/components/show-with-dependents/assembly-show-with-dependents.png
labels: [assembly, components, show]
---
{% include img.html src="assembly-show-with-dependents.png" width=250 alt="Show With Dependents command in assembly" align="center" %}

This example demonstrates how to call the 'Show With Dependents' command for components or assembly to show all components at once using SOLIDWORKS API and Windows API.

Macro will call the command for the selected component or for the assembly (if no components selected).

{% include_relative Macro.vba.codesnippet %}