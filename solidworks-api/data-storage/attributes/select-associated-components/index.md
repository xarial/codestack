---
title: Select components associated with attributes using SOLIDWORKS API
caption: Select The Components Associated With Attributes On Select
description: Example attaches to the selection events of the active assembly
labels: [attribute, component, data, example, selection, solidworks api]
redirect-from:
  - /2018/03/select-components-associated-with.html
---
This example attaches to the selection SOLIDWORKS API events of the active assembly via [NewSelectionNotify](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dassemblydocevents_newselectionnotifyeventhandler.html) notification.

If the attribute is selected and there is a component associated with this attribute - this component will be selected in the tree.  

Macro will stop once the active assembly is closed.  

*Macro module*

{% code-snippet { file-name: Macro.vba } %}

*EventsListener class*

{% code-snippet { file-name: EventsListenerClass.vba } %}