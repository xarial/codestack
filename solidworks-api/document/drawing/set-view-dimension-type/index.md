---
caption: Set View Dimension Type
title: Macro to set dimension type for all views in the active SOLIDWORKS drawing
description: VBA macro which sets dimension type (projected or true) for all drawing view in the active SOLIDWORKS drawing document
image: view-dimension-type.png
---
![View dimension type](view-dimension-type.png)

This VBA macros sets the dimension type (projected or true) for all drawing views in all sheets of the active SOLIDWORKS drawing.

Set the **DIMS_TRUE** constant to **True** to set all dimension types to **True**. Set the **DIMS_TRUE** constant to **False** to set all dimension types to **Projected**

{% code-snippet { file-name: Macro.vba } %}