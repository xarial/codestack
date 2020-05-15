---
layout: article
title: Working with assembly documents using SOLIDWORKS Document Manager API
caption: Assembly
description: Collection of examples for assemblies in Document Manager API
---
Unlike regular SOLIDWORKS API, Document Manager doesn't provide a specific interface for the assembly document rather it should be managed by [ISwDMDocument](http://help.solidworks.com/2016/english/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr~SolidWorks.Interop.swdocumentmgr.ISwDMDocument.html) and [ISwDMConfiguration2](http://help.solidworks.com/2018/english/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr~SolidWorks.Interop.swdocumentmgr.ISwDMConfiguration2.html) interfaces.

Some of the methods in those interfaces are exclusive to assembly document only, e.g [ISwDMConfiguration2::GetComponents](http://help.solidworks.com/2018/english/api/swdocmgrapi/solidworks.interop.swdocumentmgr~solidworks.interop.swdocumentmgr.iswdmconfiguration2~getcomponents.html) or [ISwDMDocument8::GetComponentCount](http://help.solidworks.com/2018/english/api/swdocmgrapi/solidworks.interop.swdocumentmgr~solidworks.interop.swdocumentmgr.iswdmdocument8~getcomponentcount.html).

It is recommended to use [ISwDMDocument::FullName](http://help.solidworks.com/2018/english/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr~SolidWorks.Interop.swdocumentmgr.ISwDMDocument~FullName.html) SOLIDWORKS Document Manager API to get the full path and match its extension to .sldasm to validate that document is an assembly.

This section contains examples and macros related to working with assembly documents using Document Manager.