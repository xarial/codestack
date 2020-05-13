---
layout: article
title: Find tables from drawing by type using SOLIDWORKS API
caption: Find Tables By Type
description: Find all tables (BOM, general, revision) from drawing sheets using SOLIDWORKS API
lang: en
image: /solidworks-api/document/tables/find-tables-by-type/drawing-view-tables.png
labels: [table,drawing]
---
{% include img.html src="drawing-view-tables.png" width=250 alt="Tables in the drawing document" align="center" %}

This examples allows to find all tables by specified type from the active drawing document using SOLIDWORKS API.

It is required to specify the array of types using the Array function, where each value represents the type of the table (BOM, general, cut-list, revision, title block etc.) as defined in [swTableAnnotationType_e](http://help.solidworks.com/2017/english/api/swconst/solidworks.interop.swconst~solidworks.interop.swconst.swtableannotationtype_e.html) enumeration.

As the result array of pointer to [ITableAnnotation](http://help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ITableAnnotation.html) SOLIDWORKS API interface is returned and title of each table is output to the immediate window of VBA editor.

{% include_relative Macro.vba.codesnippet %}