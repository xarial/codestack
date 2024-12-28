---
layout: sw-tool
caption: Set Image Quality
title: Macro to set shaded and wireframe image quality settings in SOLIDWORKS document
description: VBA macro to change the value of shaded and draft quality HLR/HLV resolution and wireframe and high quality HLT/HLV resolution
image: image-quality.svg
labels: [api, performance, resselation, image quality]
group: Performance
---

This VBA macro changes the tesselation settings of **Shaded and draft quality HLR/HLV resolution** and **Wireframe and high quality HLT/HLV resolution** settings in the SOLIDWORKS documents

![Image Quality Document Options](image-quality-document-options.png){ width=500 }

This allows to change the quality of tesselation which affects the perfromance and size of the file. Lower quality improves the performance and decreases the file size, file higher quality reduces the performance and increases the file size.

Macro can be configured by specifying the percentage of the position of the slider

~~~ vb
Const SHADED_DRAFT_QUALITY_RESOLUTION As Double = 0.25 '0 for minumum, 1 for maximum, 0.5 for the middle position
Const WIREFRAME_HIGHT_QUALITY_RESOLUTION As Double = 0.25 '0 for minumum, 1 for maximum, 0.5 for the middle position
~~~

{% code-snippet { file-name: Macro.vba } %}