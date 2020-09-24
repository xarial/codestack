---
title: Insert BOM balloons into drawing view using SOLIDWORKS API
caption: Insert BOM Balloons
description: VBA macro to automatically insert BOM balloons into an existing drawing view of the current sheet using SOLIDWORKS API
image: bom-balloons.png
labels: [BOM, balloon]
---
![BOM Balloons in the component](bom-balloons.png){ width=400 }

This VBA macro demonstrates how to insert balloons for all visible components of the first drawing view in the active drawing sheet using SOLIDWORKS API.

This view must already contain the BOM table linked to it. Macro will traverse all visible components and all visible entities of the view and will attach balloon linked to Item Number to the first visible entity.

{% code-snippet { file-name: Macro.vba } %}