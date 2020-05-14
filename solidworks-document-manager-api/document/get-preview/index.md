---
layout: article
title: Extract PNG preview image from active configuration using SOLIDWORKS Document Manager API
caption: Extract PNG Preview Image From The Active Configuration
description: Example demonstrates how to extract PNG preview image from the active configuration of SOLIDWORKS assembly or part using the document manager API.
---
Example demonstrates how to extract PNG preview image from the active configuration of SOLIDWORKS assembly or part using the document manager API.

This approach would work for both in-process and out-of-process application.

* Create C# Console application
* Paste the code
* Run the application with 2 arguments: full path to the input SOLIDWORKS part or assembly and full path to output PNG image

This example is using the [ISwDMConfiguration9::GetPreviewPNGBitmapBytes](http://help.solidworks.com/2018/english/api/swdocmgrapi/solidworks.interop.swdocumentmgr~solidworks.interop.swdocumentmgr.iswdmconfiguration9~getpreviewpngbitmapbytes.html) SOLIDWORKS Document Manager API to extract byte buffer of preview and convert it to an [Image](https://docs.microsoft.com/en-us/dotnet/api/system.drawing.image?view=netframework-4.7.2) object.

{% include_relative console.cs.codesnippet %}
