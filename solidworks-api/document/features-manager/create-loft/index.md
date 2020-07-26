---
caption: Create Loft Feature
title: Create loft feature through selected sketches or curves feature using SOLIDWORKS API
description: VBA macro to create solid loft feature from selected sketch or curve features using SOLIDWORKS API
image: loft-feature-through-curves.png
---
![Loft feature through curves](loft-feature-through-curves.png){ width=400 }

This VBA macro demonstrates how to utilize [IFeatureManager::InsertProtrusionBlend2](https://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IFeatureManager~InsertProtrusionBlend2.html) API to create loft feature from the selected sketches or curves features selected in the Feature Manager Tree.

{% code-snippet { file-name: Macro.vba } %}
