---
caption: Create Multi-Body Flat Pattern View
title: VBA macro to create flat pattern drawing view form the multi-body sheet metal part
description: VBA macro demonstrates how to create flat pattern drawing view of the multi-body sheet metal part using SOLIDWORKS API
---

This VBA example demonstrates how to create flat pattern view of a selected body from the multi-body sheet metal part.

When performing this operation manually from SOLIDWORKS, it is required to insert a drawing view of the full part, then select the single sheet metal body and set the view to **Flat Pattern**. In order to produce similar result from the API, different steps need to be performed. It is required to select the body from the visible source document before calling the [IDrawingDoc::CreateFlatPatternViewFromModelView3](https://help.solidworks.com/2013/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.idrawingdoc~createflatpatternviewfrommodelview3.html) API method.

{% code-snippet { file-name: Macro.vba } %}