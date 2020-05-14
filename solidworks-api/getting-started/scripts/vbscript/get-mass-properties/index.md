---
layout: article
title: Script extract mass properties of file using SOLIDWORKS API
caption: Get Mass Properties
description: Example demonstrates how to extract mass properties form the specified file using vbScript and SOLIDWORKS API
image: /solidworks-api/getting-started/scripts/vbscript/get-mass-properties/msgbox-mass-properties.png
labels: [mass properties, vbscript]
---
This example demonstrates how to extract mass properties from the specified file using vbScript via SOLIDWORKS API.

* Create a text file and name it as *get-mass-prps.vbs*
* Copy-paste the following code into the file

{% code-snippet { file-name: script.vbs } %}

* Save the file
* Double click to run the script
* Specify the full path to a SOLIDWORKS file (part or assembly) in the displayed input box
* As the result the following message box is displayed with mass property values

![Mass properties of the specified model are displayed in the message box](msgbox-mass-properties.png){ width=250 }
