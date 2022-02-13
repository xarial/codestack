---
caption: Get Corresponding Entity Of Derived Part
title: Get corresponding entities (faces, edges and vertices) in the derived part using SOLIDWORKS API
description: VBA macro demonstrates how to find the corresponding entities from the input part in the derived part using SOLIDWORKS API
---
[IPartDoc::InsertPart3](https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IPartDoc~InsertPart3.html) API allows to insert a derived part into another part. However the API to find the corresponding entity of the input part, similarly to [components](/solidworks-api/document/assembly/context#converting-the-pointers) is not available.

This VBA macro demonstrates a performance efficient workaround for this limitation.

## Running the macro

* Open the source part (this is the part to be inserted into another part). This part must be saved on the disc
* Select one or many entities (faces, edges, vertices). These can be selected in different bodies in case of the multi-body part
* Run the macro. Macro will index inputs and stop the execution
* Open or create new part where the source part needs to be inserted
* Continue macro execution
* As the result derived part is inserted and all the corresponding entities are selected

{% code-snippet { file-name: Macro.vba } %}
