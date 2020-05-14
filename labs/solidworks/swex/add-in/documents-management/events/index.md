---
layout: article
title: Handling the common events of SOLIDWORKS file using SwEx.AddIn framework
caption: Common Events
description: Handling of common events (rebuild, selection, configuration change, item modification, custom property modification etc.) using documents management functionality in SwEx.AddIn Framework
image: /images/codestack-snippet.png
toc_group_name: labs-solidworks-swex
labels: [events,rebuild,selection]
---
SwEx.AddIn framework exposes the common events via generic [DocumentHandler](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Core_DocumentHandler.htm):

* Save
* Selection
* Access3rdPartyData
* CustomPropertyModify
* ItemModify
* ConfigurationChange
* Rebuild
* Dimension Change

Call the [ISwAddInEx.CreateDocumentsHandler](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Base_ISwAddInEx_CreateDocumentsHandler.htm)  to create a generic events handler.

It is recommended to use the [HandleCreated](https://docs.codestack.net/swex/add-in/html/E_CodeStack_SwEx_AddIn_Base_IDocumentsHandler_1_HandlerCreated.htm) notification which will notify that new document is loaded to subscribe to the document events.

Unsubscribe from the events from [Destroyed](https://docs.codestack.net/swex/add-in/html/E_CodeStack_SwEx_AddIn_Core_DocumentHandler_Destroyed.htm) notification.

{% include code-tabs.html src="EventsAddIn" %}

Event handlers provide additional information about event, such as is it a pre or post notification and any additional parameters. Explore API reference for more information about the passed parameters.

{% include code-tabs.html src="EventsAddIn.EventHandlers" %}