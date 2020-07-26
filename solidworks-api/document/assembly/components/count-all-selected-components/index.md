---
layout: sw-tool
title: count all selected components using SOLIDWORKS API
caption: Count All Selected Components
description: Macro counts all unique components selected in the assembly and displays the result in the commands bar
image: status-bar-selected-comps.png
labels: [assembly, count components, solidworks api, status bar, utility]
group: Assembly
redirect-from:
  - /2018/03/solidworks-api-assembly-count-selected-components.html
  - /solidworks-api/document/assembly/count-all-selected-components
---
This macro counts all unique components selected in the assembly using SOLIDWORKS API. Components can be either selected in the features manager tree or in the graphics area.

Macro will also count component if only entity of the component is selected (e.g. face or edge) using [ISelectionMgr](https://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISelectionMgr.html) SOLIDWORKS API Interface..

![Quantity of selected components displayed in the status bar](status-bar-selected-comps.png){ width=320 }

{% code-snippet { file-name: Macro.vba } %}
