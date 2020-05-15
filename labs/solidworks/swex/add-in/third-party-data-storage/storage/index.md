---
layout: article
title: Storing data in the 3rd party storage store via SwEx.AddIn framework
caption: Storage
description: Serializing custom structures into the 3rd party storage store using SwEx.AddIn framework
toc-group-name: labs-solidworks-swex
order: 2
---
Call [IModelDoc2::Access3rdPartyStorageStore ](https://docs.codestack.net/swex/add-in/html/M_SolidWorks_Interop_sldworks_ModelDocExtension_Access3rdPartyStorageStore.htm) extension method to access the 3rd storage store. Pass the boolean parameter to read or write storage.

Use this approach when it is required to store multiple data structures which need to be accessed and managed independently. Prefer this instead of creating multiple [streams](/labs/solidworks/swex/add-in/third-party-data-storage/stream/)

## Storage Access Handler

To simplify the handling of the storage lifecycle, use the Documents Manager API from the SwEx.AddIn framework:

{% code-snippet { file-name: ThirdPartyData.StorageHandler.* } %}

## Reading data

[IThirdPartyStoreHandler::Storage](https://docs.codestack.net/swex/add-in/html/P_CodeStack_SwEx_AddIn_Base_IThirdPartyStoreHandler_Storage.htm) property returns null for the storage which not exists on reading.

{% code-snippet { file-name: ThirdPartyData.Storage.StorageLoad.* } %}

## Writing data

[IThirdPartyStoreHandler::Storage](https://docs.codestack.net/swex/add-in/html/P_CodeStack_SwEx_AddIn_Base_IThirdPartyStoreHandler_Storage.htm) will always return the pointer to the storage (stream is automatically created if it doesn't exist).

{% code-snippet { file-name: ThirdPartyData.Storage.StorageSave.* } %}

Explore the methods of [IComStorage](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Base_IComStorage.htm) for information of how to create sub streams or sub storages and enumerate the existing elements.