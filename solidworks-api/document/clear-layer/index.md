---
layout: sw-tool
caption: Clear Layer
title: Remove all items from the layer in SOLIDWORKS model
description: VBA macro to remove all items (annotations, sketch segments, blocks etc) from the specified layer in SOLIDWORKS document
image: remove-layer-items.svg
group: Model
---
![SOLIDWORKS layers](solidworks-layers.png)

This VBA macro collects and removes all items on the specified layer (annotations, sketch segments, blocks, sketch points and hatch). Layer itself is not removed.

Set the name of the layer in **LAYER_NAME** constant.

{% code-snippet { file-name: Macro.vba } %}