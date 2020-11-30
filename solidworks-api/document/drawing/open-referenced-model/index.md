---
layout: sw-tool
title: Open Drawing View Referenced Document
caption: VBA macro to open referenced document of the drawing view
description: VBA macro opens the document referenced by the selected drawing view in the referenced configuration and display state
image: ref-doc-display-state.svg
labels: [drawing,reference,display state]
group: Drawing
---
This VBA macro performs similar operation to **Open assembly command** on the selected SOLIDWORKS drawing view, but also activates the referenced display state associated with the drawing view.

![Open assembly command](open-assembly-command.png)

{% code-snippet { file-name: Macro.vba } %}
