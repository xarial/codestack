---
layout: article
title: Managing SOLIDWORKS documents life cycle via SwEx.AddIn framework
caption: Documents Management
description: Framework to manage SOLIDWORKS documents life cycle (open, close, activate) and its events in SwEx.AddIn
lang: en
toc_group_name: labs-solidworks-swex
order: 3
---
SwEx.AddIn frameworks provides utility class to manage document life cycle by creating a specified instance handler as a wrapper of a model.

Call [ISwAddInEx.CreateDocumentsHandler](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Base_ISwAddInEx_CreateDocumentsHandler__1.htm) method and pass the type of document handler as a generic argument or use a second overload to create a generic document handler which exposes [common events](events/) (e.g. saving, selection, rebuilding, [3rd party storage access](/labs/solidworks/swex/add-in/third-party-data-storage/)).

{% include code-tabs.html src="DocMgrAddIn.DocHandlerInit" %}

Define the document handler either by implementing the [IDocumentHandler](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Base_IDocumentHandler.htm) interface or [DocumentHandler](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Core_DocumentHandler.htm) class. 

{% include code-tabs.html src="DocMgrAddIn.DocHandlerDefinition" %}

Override methods of document handler and implement required functionality attached for each specific SOLIDWORKS model (such as handle events, load, write data etc.)

Framework will automatically dispose the handler. Unsubscribe from the custom events within the [Dispose](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Core_DocumentHandler_Dispose.htm) or [OnDestroy](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Core_DocumentHandler_OnDestroy.htm) method. The pointer to the document attached to the handler is assigned to [Model](https://docs.codestack.net/swex/add-in/html/P_CodeStack_SwEx_AddIn_Core_DocumentHandler_Model.htm) property.