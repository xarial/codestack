---
title: Replace references in components or drawing views using SOLIDWORKS Document Manager API
caption: Replace References
description: Example demonstrates how to replace references (components or drawing views) in SOLIDWORKS files using Document Manager API
labels: [document manager, references, replace, components, drawing views]
---
This example demonstrates how to replace references (components or drawing views) in SOLIDWORKS files (assemblies or drawings) using Document Manager API.

* Specify the full path to the parent file (e.g. assembly)
* Specify the full path to the document to replace
* Specify the full path to the new document

[ISwDMDocument::ReplaceReference](https://help.solidworks.com/2018/english/api/swdocmgrapi/solidworks.interop.swdocumentmgr~solidworks.interop.swdocumentmgr.iswdmdocument~replacereference.html) SOLIDWORKS Document Manager API method is used to replace the old reference with new one.

{% code-snippet { file-name: Macro.vba } %}