---
layout: article
title: VBA Macro calls Show All Components command from SOLIDWORKS API
caption: Show All Components (Show With Dependents)
description: Example demonstrates how to call the Show With Dependents command for components or assembly using SOLIDWORKS API
image: assembly-show-with-dependents.png
labels: [assembly, components, show]
---
![Show With Dependents command in assembly](assembly-show-with-dependents.png){ width=250 }

This example demonstrates how to call the 'Show With Dependents' command for components or assembly to show all components at once using SOLIDWORKS API and Windows API.

Macro will call the command for the selected component or for the assembly (if no components selected).

{% code-snippet { file-name: Macro.vba } %}