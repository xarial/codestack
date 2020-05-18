---
layout: article
title: Managing SOLIDWORKS documents via API
caption: Documents
description: Examples of closing, opening, traversing documents using SOLIDWORKS API
labels: [documents]
---
SOLIDWORKS document represented as [IModelDoc2](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2.html) interface in SOLIDWORKS API.

SOLIDWORKS allows to open and keep active multiple documents at a time. Furthermore documents might have embedded documents, for example assemblies often have another assemblies or parts linked to them as components, drawings refer underlying documents for loading drawing views and also parts can have embedded parts.

Note that the documents can be invisible (for example loaded in assembly) but still loaded into the memory and can be traversed and accessed from API methods.

This section contains code examples and macros demonstrating the techniques of managing the documents (enumerating, closing, activating, opening, identifying the types) using SOLIDWORKS API.