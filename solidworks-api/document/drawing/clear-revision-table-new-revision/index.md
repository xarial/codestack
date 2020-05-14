---
layout: article
title: Clear revision table and add new revision using SOLIDWORKS API
caption: Clear Revision Table And Add New Revision
description: Example finds the revision table and removes all revisions and then adds new row with custom data
image: /solidworks-api/document/drawing/clear-revision-table-new-revision/sw-revision-table.png
labels: [add revision, clear revisions, drawing.revision table, example, solidworks api]
redirect-from:
  - /2018/03/solidworks-api-drawing-clear-rev-table-add-new-row.html
---
This example finds the revision table and removes all revisions and then adds new row with custom data using SOLIDWORKS API.

![Revision Table](sw-revision-table.png){ width=640 }

[IRevisionTableAnnotation](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.irevisiontableannotation.html) SOLIDWORKS API interface is used to manage specific functionality of this type of the table.

{% code-snippet { file-name: Macro.vba } %}
