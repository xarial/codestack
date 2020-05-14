---
layout: article
title: Importing and exporting files using SOLIDWORKS API
caption: Import/Export
description: Collection of examples and articles related to files importing and exporting using SOLIDWORKS API
image: /solidworks-api/import-export/import-export.img
order: 7
---
{% include img.html src="import-export.svg" width=250 alt="Importing and exporting SOLIDWORKS files with SOLIDWORKS API" align="center" %}

SOLIDWORKS API allows to export file to foreign format via [IModelDocExtension::SaveAs2](http://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldocextension~saveas2.html) and it is only required to specify the corresponding extension for the foreign file (e.g. stp, x_t, igs etc.).

For importing foreign file it is required to use the [ISldWorks::LoadFile4](http://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~loadfile4.html) SOLIDWORKS API method. Parameters of this function allows to specify additional importing options.

This section contains code example, macros and scripts for automation import and export functionality of SOLIDWORKS.
