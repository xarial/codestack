---
caption: Traverse Loop Points
title: VBA macro to traverse vertices of loops of the selected SOLIDWORKS face
description: VBA macro to traverse all vertices of inner and outer loops of the selected face of active SOLIDWORKS document
---

This VBA macro traverses outer loop in CCW direction and innter loops in the CW direction for the selected face in the active SOLIDWORKS document.

Macro selects vetices one-by-one and pauses and execution on each selection

{% code-snippet { file-name: Macro.vba } %}