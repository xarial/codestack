---
caption: Combine Identical Components
title: Combine identical components command in SOLIDWORKS BOM table
description: Macro to emulate combine identical components command in SOLIDWORKS BOM table
image: combine-identical-components.png
---
![Combine identical components command](combine-identical-components.png)

This VBA macro demonstrates how to emulate the *Combine identical component* command which is missing in SOLIDWORKS API.

Select BOM table to combine identical components. By default, all components are combined, however it is possible to specify the rows to combine by changing the parameters of **CombineIdenticalComponents** function in the macro.

{% code-snippet { file-name: Macro.vba } %}