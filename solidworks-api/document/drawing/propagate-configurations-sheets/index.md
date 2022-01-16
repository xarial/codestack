---
caption: Propagate Configurations To Sheets
title: Macro propagates configurations of the referenced document to sheets in the SOLIDWORKS drawings
description: VBA macro copies the input sheet and sets the referenced configuration sof the referenced document
image: sheets.png
---
![Drawings with multiple sheets](sheets.png){ width=800 }

This VBA macro will copy the active sheet and propagate referenced configurations to each copy.

Macro will automatically set the referenced configuration on each new sheet and rename the sheet based on the configuration name.

As teh result drawing will contain sheets for all the configurations of the multi-body part or assembly.

{% code-snippet { file-name: Macro.vba } %}