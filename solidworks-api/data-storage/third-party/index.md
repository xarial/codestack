---
layout: article
title: Data saving in the 3rd party storage using SOLIDWORKS API
caption: 3rd Party Storage And Store
description: Section explaining how to use 3rd party storage and 3rd party store in SOLIDWORKS API to serialize and deserialize the data directly in the model stream
lang: en
image: /solidworks-api/data-storage/third-party/store-diagram.png
labels: [store,3rd party,third party,storage,serialization]
---
3rd party storage and 3rd party store are the containers for the external applications (add-ins, macros, stand alone applications) to store serialize the data directly in the model stream.

This technique allows to store the complex data and provides best performance options to read and write large amount of data.

SOLIDWORKS enables to store the data in 2 different containers:

* Storage (Stream)
* Storage Store

If File System is taken as analogue the Storage would correspond to file while Storage Store to folder. Storage Stores can have sub streams or sub stores.

The following diagram explains the structure of the SOLIDWORKS model storages. Red elements represent the containers managed directly by SOLIDWORKS while other elements represent the containers managed by 3rd parties.

{% include img.html src="store-diagram.svg" width=550 alt="Document Store Diagram" align="center" %}

## 3rd Party Storage

This is a container which is managed via [IStream](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-istream) interface. This option is used when application only needs to store the single data structure (e.g. XML tree, text, image, binary data).

In order to get the pointer to the stream (both for reading or writing) the [IModelDoc2::IGet3rdPartyStorage](http://help.solidworks.com/2015/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~IGet3rdPartyStorage.html) SOLIDWORKS API method should be called and corresponding flag is passed.

### Notes

* If stream was never written before the [IModelDoc2::IGet3rdPartyStorage](http://help.solidworks.com/2015/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~IGet3rdPartyStorage.html) method returns null.
* Stream should always be released after the get method called via [IModelDoc2::IRelease3rdPartyStorage](http://help.solidworks.com/2015/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~IRelease3rdPartyStorage.html). This also applies when get method returns null (i.e. stream was not stored before)
* [IStream::Commit](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nf-objidl-istream-commit) method should not be called when storing the data otherwise *Method Not Implemented* exception will be thrown.

### Lifecycle

Storage is available for reading between the [LoadFromStorage](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dpartdocevents_loadfromstoragenotifyeventhandler.html) notification and the destroying of the model. LoadFromStorageStore available for [part](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dpartdocevents_loadfromstoragenotifyeventhandler.html), [assembly](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dassemblydocevents_loadfromstoragenotifyeventhandler.html)  and [drawing](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ddrawingdocevents_loadfromstoragenotifyeventhandler.html) 

Storage is available for writing only within the [SaveToStorage](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dpartdocevents_savetostoragenotifyeventhandler.html) notification for [part](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dpartdocevents_savetostoragenotifyeventhandler.html), [assembly](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dassemblydocevents_savetostoragenotifyeventhandler.html) and [drawing](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ddrawingdocevents_savetostoragenotifyeventhandler.html) correspondingly.

## 3rd Party Storage Store

This is a container which is managed via [IStorage](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-istorage) interface. This option is used when application manages complex sets of data and access to certain portions is required at certain times. Storage container allows to create sub streams and sub storages to manage the data and only specific streams can be accessed when required avoiding the need to load the whole structure into the memory.

To get the pointer to the storage the [IModelDocExtension::IGet3rdPartyStorageStore](http://help.solidworks.com/2015/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~IGet3rdPartyStorageStore.html) SOLIDWORKS API method needs to be called.

### Notes
* [IModelDocExtension::IGet3rdPartyStorageStore](http://help.solidworks.com/2015/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~IGet3rdPartyStorageStore.html) returns null for the storage which was never written before
* Similar to streams, store always needs to be released via [IModelDocExtension::IRelease3rdPartyStorageStore](http://help.solidworks.com/2015/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~IRelease3rdPartyStorageStore.html) method.
* Use methods of [IStorage](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-istorage) interface to create sub streams and storages.

### Lifecycle

Storage is available for reading between the [LoadFromStorageStore](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dpartdocevents_loadfromstoragestorenotifyeventhandler.html) notification and the destroying of the model. LoadFromStorageStore available for [part](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dpartdocevents_loadfromstoragestorenotifyeventhandler.html), [assembly](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dassemblydocevents_loadfromstoragestorenotifyeventhandler.html)  and [drawing](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ddrawingdocevents_loadfromstoragestorenotifyeventhandler.html) 

Storage is available for writing only within the [SaveToStorageStore](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dpartdocevents_savetostoragestorenotifyeventhandler.html) notification for [part](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dpartdocevents_savetostoragestorenotifyeventhandler.html), [assembly](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dassemblydocevents_savetostoragestorenotifyeventhandler.html) and [drawing](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ddrawingdocevents_savetostoragestorenotifyeventhandler.html) correspondingly.

## Usage

Usually 3rd party containers (storage and store) are used in add-ins when model is complemented with additional functionality (e.g. electrical data, PDM, security, etc.). In this case this additional information is usually displayed in the Feature Tree, Task Panes etc. and loaded when model is opened and saved together with the model making this approach a fully integrated solutions.

*SaveToStorage* and *SaveToStorageStore* SOLIDWORKS API notifications are raised directly after the File Save Notification which means that there is no need to implement custom saving of the data as it will be automatically triggered via user saving.

The best place to attach save and load event would be within the [DocumentLoadNotify](http://help.solidworks.com/2015/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.dsldworksevents_documentloadnotify2eventhandler.html) event.

When 3rd party data is modified (e.g. user added new node in the 3rd party tree) it is recommended to mark model as dirty via [IModelDoc2::SetSaveFlag](http://help.solidworks.com/2015/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~SetSaveFlag.html) which indicates that model required to be saved by the user.

## Storage And Streams Naming Conflicts

Storages and stores accessed by the corresponding names. It might be the cases when different developers might use the same name for storage or store. In this case conflict occurs. When using 3rd party containers it is recommended to register the storage or store name via SOLIDWORKS API Support and in this case this name will be reserved.

Refer [Storing 3rd party data in SOLIDWORKS models using SwEx.AddIn framework]({{ "/labs/solidworks/swex/add-in/third-party-data-storage/" | relative_url }}) article for the information of how to access 3rd party containers using SwEx.AddIn framework.