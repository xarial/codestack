---
layout: article
title: Storing data in the 3rd party storage (stream) via SwEx.AddIn framework
caption: Stream
description: Serializing custom structures into the 3rd party storage (stream) using SwEx.AddIn framework
toc-group-name: labs-solidworks-swex
order: 1
---
Call [IModelDoc2::Access3rdPartyStream](https://docs.codestack.net/swex/add-in/html/M_SolidWorks_Interop_sldworks_ModelDocExtension_Access3rdPartyStream.htm) extension method to access the 3rd party stream. Pass the boolean parameter to read or write stream.

use this approach when it is required to store a single structure at the model.

## Stream Access Handler

To simplify the handling of the stream lifecycle, use the Documents Manager API from the SwEx.AddIn framework:

{#% include code-tabs.html src="ThirdPartyData.StreamHandler" %}

## Reading data

[IThirdPartyStreamHandler::Stream](https://docs.codestack.net/swex/add-in/html/P_CodeStack_SwEx_AddIn_Base_IThirdPartyStreamHandler_Stream.htm) property returns null for the stream which not exists on reading.

{#% include code-tabs.html src="ThirdPartyData.Stream.StreamLoad" %}

## Writing data

[IThirdPartyStreamHandler::Stream](https://docs.codestack.net/swex/add-in/html/P_CodeStack_SwEx_AddIn_Base_IThirdPartyStreamHandler_Stream.htm) will always return the pointer to the stream (stream is automatically created if it doesn't exist).

{#% include code-tabs.html src="ThirdPartyData.Stream.StreamSave" %}
