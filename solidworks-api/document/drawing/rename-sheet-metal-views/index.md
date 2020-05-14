---
layout: sw-tool
title: Rename flat pattern views with cut-list names VBA macro
caption: Rename Flat Pattern Views With Cut-List Names
description: VBA macro to rename all flat pattern views in the the active sheet after the respective cut-list names using SOLIDWORKS API
image: /solidworks-api/document/drawing/rename-sheet-metal-views/renamed-flat-pattern-drawing-view.png
labels: [rename view,cut list,flat pattern]
group: Drawing
---
![Cut-list for sheet metal body](cut-list-name.png){ width=250 }

Cut list names for sheet metal bodies can be used to store important information, such as part number. This VBA macro allows to rename all flat pattern views of sheet metal in the active drawing sheet with the name of the respective cut-list item using SOLIDWORKS API.

![Drawing view renamed after the cut-list](renamed-flat-pattern-drawing-view.png){ width=250 }

{% code-snippet { file-name: Macro.vba } %}
