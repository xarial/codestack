---
layout: article
title: Using persistent reference id in SOLIDWORKS API to track objects
caption: Persistent Reference Id
description: This article explains the use of persistent reference ids to track any selectable entity across SOLIDWORKS sessions
image: /solidworks-api/document/tracking-objects/persist-references/persist-id-array.png
labels: [persistent, reference, tracking]
---
Persistent reference ids available in SOLIDWORKS API allow to retrieve the persistent link to any selectable object in SOLIDWORKS. The main benefit of persistent reference is its life cycle as the reference remains valid across rebuild operations, SOLIDWORKS sessions or even SOLIDWORKS releases.

Persistent reference id is an array of bytes. This array may change for the same reference so it is not possible to compare two arrays to identify if the references are the same. Use [IModelDocExtension::IsSamePersistentID method](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldocextension~issamepersistentid.html) to identify if two persist references are the same.

{% include img.html src="persist-id-array.png" width=350 alt="Array of bytes of persist reference displayed in the watch window of VBA Editor" align="center" %}

Even if array may change for the same entity it is still possible to retrieve the valid pointer to the entity via [IModelDocExtension::GetPersistReference3](http://help.solidworks.com/2017/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldocextension~getpersistreference3.html) SOLIDWORKS API method.

The following example outputs the persist id of any selected entity into immediate window in the format of base64 string

{% include img.html src="immediate-window-persist-id.png" alt="Persist reference id converted to base64 string displayed in the immediate window of VBA Editor" align="center" %}

Use this example to read the id of the entity.

> The id output to immediate Window might contain line break. It should be removed from the id and should be considered as single line string

{% code-snippet { file-name: get-id.vba } %}

The following example allows to select the object by retrieving its pointer from persist id.

* Copy the id from the previous macro
* Close the sample model
* Reopen the model and run the example.
* Enter the copied id into the box
* The entities selected in previous example is re-selected

{% code-snippet { file-name: get-object.vba } %}
