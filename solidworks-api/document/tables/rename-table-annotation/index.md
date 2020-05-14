---
layout: article
title: Macro renames table annotation using SOLIDWORKS API
caption: Rename Table Annotation
description: Example demonstrates how to rename the selected table using SOLIDWORKS API
image: /solidworks-api/document/tables/rename-table-annotation/rename-table-annotation.png
labels: [table, rename]
---
{% include img.html src="rename-table-annotation.png" width=450 alt="Table annotation renamed to a custom name" align="center" %}

This example demonstrates how to rename the selected table using SOLIDWORKS API via [ITableAnnotation](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ITableAnnotation.html) interface. Table should be selected in the graphics area (not in the feature tree)

Specify the name of the table by modifying the constant at the beginning of the macro:

~~~ vb
Const TABLE_NAME As String = "MyTable"
~~~

{% code-snippet { file-name: Macro.vba } %}
