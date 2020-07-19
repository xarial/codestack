---
layout: sw-tool
title: Macro to export selected bodies to foreign format
caption: Export Selected Bodies
description: VBA macro to export only selected bodies to foreign format (e.g. 3D xml, xaml, amf, 3mf)
image: export-bodies.svg
labels: [export, bodies]
group: Import/Export
---
When exporting part file to most of the foreign format supported by SOLIDWORKS it is possible to select the scope bodies of export, allowing to only process selected bodies.

![Export bodies dialog](export-dialog.png)

However this feature is not supported by all formats. For example the formats such as 3D xml, xaml, amf, 3mf will always export all bodies, regardless of the selection.

This VBA macro allows to export only selected bodies to any format supported by SOLIDWORKS.

Select the bodies, faces, edges or vertices and run the macro and specify the name of export to produce a result.

{% code-snippet { file-name: Macro.vba } %}
