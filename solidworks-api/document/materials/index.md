---
layout: article
title: Working with part materials using SOLIDWORKS API
caption: Materials
description: Collection of articles and examples related to materials handling using SOLIDWORKS API
lang: en
order: 15
image: /images/codestack-snippet.png
---
Material database in SOLIDWORKS is stored within the XML file. SOLIDWORKS API doesn't provide any direct methods of working with materials database (i.e. searching, adding, reading etc.). However as XML is an open format any XML parsing techniques would apply, i.e. using the [XmlDocument](https://docs.microsoft.com/en-us/dotnet/api/system.xml.xmldocument), [XmlSerializer](https://docs.microsoft.com/en-us/dotnet/api/system.xml.serialization.xmlserializer) etc.

[ISldWorks::GetMaterialDatabases](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~getmaterialdatabases.html) SOLIDWORKS API method returns the paths to material databases.

[IPartDoc::GetMaterialPropertyName2](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ipartdoc~getmaterialpropertyname2.html) returns the name of the material and the name of the database this material is stored in.

This section contains examples explaining how to work with the materials database in SOLIDWORKS, how to apply and read material information from SOLIDWOKS parts and bodies.