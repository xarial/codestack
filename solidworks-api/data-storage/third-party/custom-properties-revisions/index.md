---
layout: article
title: Save custom properties revisions into 3rd party storage store using SOLIDWORKS API
caption: Save Custom Properties Revisions
description: Saving custom properties revisions (snapshots) into document 3rd party storage store using SOLIDWORKS API
lang: en
image: /solidworks-api/data-storage/third-party/custom-properties-revisions/properties-snapshots-data.png
labels: [storage,3rd party,custom properties]
---
{% include img.html src="custom-properties.png" width=450 alt="Custom Properties" align="center" %}

This example demonstrates how to utilize 3rd party storage store to save file custom properties revisions using SOLIDWORKS API.

This add-in is built using the [SwEx.AddIn]({{ "/labs/solidworks/swex/add-in/" | relative_url }}) framework but it could work with any other methods of creating the add-ins.

Add-in adds two buttons in the menu and toolbar and provides two handlers correspondingly: 

* TakeCustomPropertiesSnapshot - reads current state of custom properties and serializes it to the 3rd party storage
* LoadSnapshots - loads all revisions and displays the message box

The snapshot of each revision is stored within the storages (streams) in 3rd party sub store, while information about all available snapshots is saved into the sub stream of 3rd party storage.

## Usage Instructions

* Open any existing SOLIDWORKS models (part, assembly or drawing)
* Add some custom properties into *Custom* tab
* Click *TakeCustomPropertiesSnapshot* from the *Tools\Custom Properties Revisions* menu
* Modify properties and click *TakeCustomPropertiesSnapshot* again. Repeat if needed
* You can close and reopen the model and SOLIDWORKS. Click *LoadSnapshots* command. All properties revisions are displayed in the message box

{% include img.html src="properties-snapshots-data.png" width=450 alt="All properties revisions displayed in the message box" align="center" %}

### PropertiesSnapshot.cs

Structures to serialize properties and info

{% include_relative PropertiesSnapshot.cs.codesnippet %}

### CustomPropertiesRevisionsAddIn.cs

Add-in class which is handling the menu commands and reads and outputs the data

{% include_relative CustomPropertiesRevisionsAddIn.cs.codesnippet %}

### CustomPropertiesRevisions.cs

Functions to access storage and store to serialize and deserialize the data

{% include_relative CustomPropertiesRevisions.cs.codesnippet %}

### ComStorage.cs

Wrapper around [IStorage](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-istorage) interface which simplifies the access from .NET language

{% include_relative ComStorage.cs.codesnippet %}

### ComStream.cs

Wrapper around [IStream](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-istream) interface which simplifies the access from .NET language

{% include_relative ComStream.cs.codesnippet %}
