---
layout: article
title: Utilizing main ISwDMApplication application object in SOLIDWORKS Document Manager API
caption: Application
description: Explanation and examples of top level object in Document Manager API ISwDMApplication
image: /images/codestack-snippet.png
---
[ISwDMApplication](http://help.solidworks.com/2017/english/api/swdocmgrapi/solidworks.interop.swdocumentmgr~solidworks.interop.swdocumentmgr.iswdmapplication.html) is a top level object in SOLIDWORKS Document Manager API hierarchy and represents the application itself.

Pointer to the object can be accessed via [ISwDMClassFactory::GetApplication](http://help.solidworks.com/2017/english/api/swdocmgrapi/SOLIDWORKS.Interop.swdocumentmgr~SOLIDWORKS.Interop.swdocumentmgr.ISwDMClassFactory~GetApplication.html) method.

### Functionality

* Accessing the documents (i.e. opening the document stream)
* Operations with documents (moving, copying) with an ability to preserver references
* Creating the data objects (such as search options or external reference options)