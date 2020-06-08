---
title: Get all external references of document using SOLIDWORKS Document Manager API
caption: Get All External References
description: Macro demonstrates how to extract all external references (including nested references) for specified SOLIDWORKS file using Document Manager API
labels: [document manager, external references, components]
---
This macro demonstrates how to extract all external references (including nested references, assembly components, drawing views) for specified SOLIDWORKS file (part, assembly or drawing) using SOLIDWORKS Document Manager API.

Modify the macro and specify the full path to the root file to collect references from.

Run the macro. All references are output to the immediate window.

Macro is using the [ISwDMDocument21::GetAllExternalReferences5](http://help.solidworks.com/2018/english/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr~SolidWorks.Interop.swdocumentmgr.ISwDMDocument21~GetAllExternalReferences5.html) SOLIDWORKS Document Manager API to list all the dependencies of the files. This method is called recursively to collect the references at all levels of SOLIDWORKS assembly.

{% code-snippet { file-name: Macro.vba } %}