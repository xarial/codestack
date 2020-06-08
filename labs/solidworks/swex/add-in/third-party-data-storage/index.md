---
title: Storing 3rd party data in SOLIDWORKS models using SwEx.AddIn framework
caption: 3rd Party Data Storage
description: Storing 3rd party data (structures and storages) in SOLIDWORKS model stream using SwEx.AddIn framework
toc-group-name: labs-solidworks-swex
order: 5
---
SwEx.AddIn framework provides functionality for handling the data (both serializing and deserializing) in SOLIDWORKS 3rd party [storage (stream)](stream) and [storage store](storage).

Refer [Data saving in the 3rd party storage using SOLIDWORKS API](/solidworks-api/data-storage/third-party/) for detailed overview of 3rd party storage and store.

It is recommended to use this functionality in conjunction with [Documents Management](/labs/solidworks/swex/add-in/documents-management/) by overriding the [OnLoadFromStream](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Core_DocumentHandler_OnLoadFromStream.htm), [OnSaveToStream](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Core_DocumentHandler_OnSaveToStream.htm), [OnLoadFromStorageStore](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Core_DocumentHandler_OnLoadFromStorageStore.htm), [OnSaveToStorageStore](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Core_DocumentHandler_OnSaveToStorageStore.htm) methods.

See the short video demonstration of usage 3rd party storage using SwEx.AddIn framework:

{% youtube { id: 9Y_OsoauvuQ } %}