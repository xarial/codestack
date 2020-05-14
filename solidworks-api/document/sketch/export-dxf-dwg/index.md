---
layout: sw-tool
title: Macro to export selected sketch in SOLIDWORKS file to DXF/DWG file
caption: Export Sketch To DXF/DWG
description: VBA macro to export the selected 2D sketch in SOLIDWORKS part or assembly file to the DXF or DWG file
image: /solidworks-api/document/sketch/export-dxf-dwg/sketch-dwf-dwg.png
logo: /solidworks-api/document/sketch/export-dxf-dwg/dxf-sketch.svg
labels: [sketch,export,dxf,dwg]
categories: sw-tools
group: Import/Export
---
{% include img.html src="sketch-dwf-dwg.png" width=350 alt="DXF/DWG file created from the sketch" align="center" %}

This VBA macro exports the selected 2D sketch in part or assembly to DXF or DWG file.

## Options

Configure the name of the output file by modifying the *EXPORT_NAME_TEMPLATE* constant as shown below using free text and placeholders.

* \[title\] placeholder will be replaced with the title of the original part or assembly file (without extension)
* \[sketch\] placeholder will be replaced with the name of the sketch DXF\DWG file created from

Specify the extension (.dxf or .dwg) in the file template

File wil be saved in the same directory as original part or assembly document.

~~~ vb
Const EXPORT_NAME_TEMPLATE As String = "ExportFile_[title]_[sketch].dxf"
~~~

{% code-snippet { file-name: Macro.vba } %}