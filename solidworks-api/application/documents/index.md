---
layout: article
title: Managing SOLIDWORKS documents via API
caption: Documents
description: Examples of closing, opening, traversing documents using SOLIDWORKS API
lang: en
image: /images/codestack-snippet.png
labels: [documents]
---
SOLIDWORKS document represented as [IModelDoc2](http://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2.html) interface in SOLIDWORKS API.

This section contains code examples and macros demonstrating the techniques of managing the documents (enumerating, closing, activating, opening, identifying the types) using SOLIDWORKS API.

Note that the documents can be invisible (for example loaded in assembly) but still loaded into the memory and can be traversed and accessed from API methods.