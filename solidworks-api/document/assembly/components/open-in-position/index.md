---
layout: sw-tool
title: Open all selected components in positions in new windows
caption: Open Components In Positions
description: VBA macro to open each selected component in the assembly in the separate window in the same position they appear in the current assembly view
image: /solidworks-api/document/assembly/components/open-in-position/open-in-position.png
logo: /solidworks-api/document/assembly/components/open-in-position/open-in-position.svg
labels: [position,component]
group: Assembly
---
This VBA macro opens all selected components in the active assembly in their own windows in the same position as they appear in the original SOLIDWORKS assembly.

This macro emulates the *Open Part In Position* command in SOLIDWORKS toolbar, but allows to open multiple selected components at the same time.

![Open part in position command](open-part-in-position-command.png){ width=250 }

{% code-snippet { file-name: Macro.vba } %}
