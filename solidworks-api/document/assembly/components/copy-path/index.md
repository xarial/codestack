---
layout: sw-tool
title: Macro to copy path of SOLIDWORKS component to clipboard
caption: Copy Component Path
description: Macro copies the path of the selected component in assembly or drawing into the clipboard using SOLIDWORKS API
image: copy-component-path.png
labels: [path, clipboard, component]
group: Assembly
---
![Component selected in the feature tree](selected-component.png){ width=250 }

This macro copies the full path to the selected component into the clipboard using SOLIDWORKS API.

* Component can be selected in assembly or drawing document
* Component can be selected in the feature tree or in the graphics area
    * It is also possible to select a component entity (i.e. face or edge) to get the path to the component

{% code-snippet { file-name: Macro.vba } %}
