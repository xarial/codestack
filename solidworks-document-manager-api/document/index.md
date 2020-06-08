---
title: Examples of using ISwDMDocument object for document in Document Manager API
caption: Document
order: 2
description: Examples related of usage of ISwDMDocument interface in document manager API
labels: [document manager, document]
---
[ISwDMDocument](http://help.solidworks.com/2016/english/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr~SolidWorks.Interop.swdocumentmgr.ISwDMDocument.html) SOLIDWORKS Document Manager API interface represents the stream for SOLIDWORKS file (part, assembly and drawing). Document can be opened for read-only access or for write access. This option is controlled by **allowReadOnly** parameter of [ISwDMApplication::GetDocument](http://help.solidworks.com/2012/english/api/swdocmgrapi/solidworks.interop.swdocumentmgr~solidworks.interop.swdocumentmgr.iswdmapplication~getdocument.html) method.

When document is opened read-only any modification won't be saved. 