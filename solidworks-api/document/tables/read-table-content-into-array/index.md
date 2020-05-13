---
layout: article
title: Read table content into array using SOLIDWORKS API
caption: Read Table Content Into Array
description: Example demonstrates how to read the content of the selected table (Bill Of Materials, General Table, Cut-List Table etc.) into the 2-dimensional array
lang: en
image: /images/codestack-snippet.png
labels: [array, bom, read, solidworks api, table]
redirect_from:
  - /2018/03/solidworks-api-model-read-table-content-into-array.html
---
This example demonstrates how to read the content of the selected table (Bill Of Materials, General Table, Cut-List Table etc.) into the 2-dimensional array using SOLIDWORKS API.

[ITableAnnotation](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ITableAnnotation.html) SOLIDWORKS API interface provides an access to the data of all table types.

{% include_relative Macro.vba.codesnippet %}