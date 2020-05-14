---
layout: article
title: Add watermark to model using 3rd party storage via SOLIDWORKS Document Manager API
caption: Add Watermark To Model
description: Add digital watermark into model using 3rd party storage via SOLIDWORKS Document Manager API
image: /solidworks-document-manager-api/document/data-storage/third-party/add-watermark/add-watermark-console-output.png
labels: [watermark, storage]
---
This example demonstrates how to add digital watermark into SOLIDWORKS model (part, assembly or drawing) into 3rd party storage (stream) using SOLIDWORKS Document Manager API.

This application implemented as the command line program and has the following arguments

## Adding watermark

* Full path to SOLIDWORKS file
* -w - Flag to indicate that watermark needs to be added
* Company Name - name of the company to add to waternark

~~~
AddWatermark.exe C:\MyPart.sldprt -w MyCompanyName
~~~

Watermark will include company name, current use name and time stamp

## Reading watermark

* Full path to SOLIDWORKS file
* -r - Flag to indicate that watermark needs to be read

~~~
AddWatermark.exe C:\MyPart.sldprt -r MyCompanyName
~~~

As the result the stored watermark is displayed in the console application

{% include img.html src="add-watermark-console-output.png" width=450 alt="Output results in the console" align="center" %}

### Program.cs

Console application containing the routing for reading and adding watermark

{% include_relative Program.cs.codesnippet %}

### ComStream.cs

Wrapper around [IStream](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-istream) interface which simplifies the access from .NET language

{% include_relative ComStream.cs.codesnippet %}
