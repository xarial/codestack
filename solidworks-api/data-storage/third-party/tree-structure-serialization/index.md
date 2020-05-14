---
layout: article
title: Tree structure serialization in model 3rd party storage using SOLIDWORKS API
caption: Third Party Store Tree Serialization
description: Example of usage of 3rd Party Storage (stream) to serialize and deserialize tree structure using SOLIDWORKS API and XmlSerializers within the model document
image: /solidworks-api/data-storage/third-party/tree-structure-serialization/read-data-result.png
labels: [serialization,third party store]
---
This example demonstrates how to use 3rd Party Storage in SOLIDWORKS API to read and write custom structure directly within the model.

Example SOLIDWORKS add-in is built using the [SwEx.AddIn](/labs/solidworks/swex/add-in/) framework but it could work with any other methods of creating the add-ins.

Add-in adds two buttons in the menu and toolbar and provides two handlers correspondingly: 

* SaveTree - asynchronous method to store the data in the stream. This method bumps the revision of the structure after each save.
* LoadTree - loads the data from the stream and displays the name of the root element and the version

![Result displayed from the data read from the stream](read-data-result.png){ width=250 }

## Usage Instructions

* Open any model
* Click "Save Data" button. First version of the structure is saved with the model
* You can close the model and SOLIDWORKS
* Reopen the model and click "Load Data". Information about saved structure is displayed in the message box
* Click "Save Data" button again. Data version is updated

It is required to set the 'Allow unsafe code' option in the Visual Studio Project settings:

![Allow unsafe code option in C# project](vs-setting-allow-unsafe-code.png){ width=450 }

**TreeSerializerAddIn.cs**

{% code-snippet { file-name: TreeSerializerAddIn.cs } %}

Structure used in this example represents the simple hierarchical data

**ElementsTree.cs**

{% code-snippet { file-name: ElementsTree.cs } %}

For simplicity [IStream](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-istream) com stream is wrapped into the [System.IO.Stream](https://docs.microsoft.com/en-us/dotnet/api/system.io.stream?view=netframework-4.7.2) type.

**ComStream.cs**

{% code-snippet { file-name: ComStream.cs } %}

Serialization and deserialization routine utilizing the [XmlSerializer](https://docs.microsoft.com/en-us/dotnet/api/system.xml.serialization.xmlserializer?view=netframework-4.7.2) class, but any other serialization methods could be used.

**TreeSerializer.cs**

{% code-snippet { file-name: TreeSerializer.cs } %}
