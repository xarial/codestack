---
caption: Scale Imported Geometry
title: VBA macro to scale the geometry of the imported features using SOLIDWORKS API
description: VBA macro scales the bodies from the imported features of the foreign formats (e.g. STEP, IGES) with the specified scale factor
image: imported-feature.png
---
![Imported geometry feature](imported-feature.png){ width=250 }

This VBA macro scales all bodies form the imported features in active SOLIDWORKS part file. THe imported features will be generated if file is loaded from neutral formats like STEP, IGES, Parasolid unless 3D Interconnect option is used.

Set the scale factor in the **SCALE_FACTOR** constant.

{% code-snippet { file-name: Macro.vba } %}