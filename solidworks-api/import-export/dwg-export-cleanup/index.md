---
layout: article
title: Export Flat Pattern to DXF/DWG with Cleanup page using SOLIDWORKS API
caption: Export Flat Pattern With Cleanup
description: VBA example which demonstrates how to export specified flat pattern to DXF/DWG with Cleanup dialog
image: /solidworks-api/import-export/dwg-export-cleanup/cleanup-page.png
labels: [dxf,dwg,cleanup,flat pattern,export]
---
[IPartDoc::ExportToDwg2](https://help.solidworks.com/2014/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IPartDoc~ExportToDWG2.html) SOLIDWORKS API method allows to export selected flat pattern to DXF/DWG format. But this API doesn't allow to show the built in Cleanup dialog to modify the DXF/DWG before exporting.

{% include img.html src="cleanup-page.png" width=350 alt="DXF/DWG Cleanup" align="center" %}

The code below provides a workaround for this issue.

> Not this code doesn't allow to set the settings of the export (default options are used). It is required to use Windows API to modify the options and check boxes.

## Configuration

Specify the macro parameters as shown below:

~~~vb
Const FLAT_PATTERN_FEAT_NAME As String = "Flat-Pattern1" 'name of flat pattern feature to export
Const OUT_FILE_NAME As String = "D:\sample.dxf" 'output file name for the export
~~~

## Macro Module

{% include_relative Macro.vba.codesnippet %}

## ExportEventsListener Class module

Create new [class module](/visual-basic/classes/) with name **ExportEventsListener** and add the code below

{% include_relative ExportEventsListener.vba.codesnippet %}
