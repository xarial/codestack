---
layout: article
title: Add invisible custom property to the model using SOLIDWORKS Document Manager API
caption: Add Invisible Custom Property
description: VBA macro to add invisible custom property to the model using SOLIDWORKS Document Manager API
lang: en
image: /images/codestack-snippet.png
labels: [invisible, custom property]
---
SOLIDWORKS models contain several invisible custom properties, such as $PRP:"SW-File Name", $PRP:"SW-Title". Those are read-only and cannot be modified from the user interface. It is however possible to add new custom property using Document Manager API. This property is not available in the custom property manager page and cannot be modified by the user or SOLIDWORKS API.

This VBA example shows how to add the invisible custom property for the specified model. Configure the macro as follows:

~~~ vb
Const FILE_PATH As String = "C:\SampleModel.SLDPRT" 'Full path to file to add invisible property to
Const PRP_NAME As String = "MyProperty" 'Property name to add
Const PRP_VAL As String = "MyValue" 'Property value to assign
~~~

{% include_relative Macro.vba.codesnippet %}