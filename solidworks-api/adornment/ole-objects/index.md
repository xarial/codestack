---
title: Managing OLE Objects in models using SOLIDWORKS API
caption: OLE Objects
description: Collection of macros and examples which demonstrate how to work with different embedded OLE objects (design table, attachment etc.) using SOLIDWORKS API
order: 2
labels: [ole, embeding]
---
Object Linking and Embedding (OLE) is a Microsoft technology allowing to inserted 3rd party application objects into the documents. In SOLIDWORKS OLE objects are used to represent Design Tables, Attachment and any file dropped directly into the Document.

Such objects usually can be manipulated directly from the host environment. For example embeded Excel file can be modified without exiting the SOLIDWORKS window.

OLE Objects are usually saved with SOLIDWORKS file and can be removed, resized or used directly in the graphics area.

SOLIDWORKS API enables the access to OLE objects via [ISwOLEObject](http://help.solidworks.com/2018/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISwOLEObject.html) interface. Objects can be enumerated, created and deleted by using the API methods of [IModelDocExtension](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension.html) interface.

This section contains macros and examples allowing to manipulate OLE objects in documents using the SOLIDWORKS API.