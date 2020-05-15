---
layout: article
title: Get and select all visible components in assembly using SOLIDWORKS API
caption: Get And Select Visible Components Only
description: VBA macro example to get and select all visible components (not suppressed and not hidden) using SOLIDWORKS API
image: components-tree.png
labels: [components,suppressed,hidden,select]
---
![Components selected in the feature manager tree](components-tree.png){ width=350 }

This VBA macro gets all the pointers to the visible (not suppressed and not hidden) components in the active assembly. All the components are selected using multi-select SOLIDWORKS API.

{% code-snippet { file-name: Macro.vba } %}