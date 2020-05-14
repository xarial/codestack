---
layout: article
title: Read All Custom Properties using SOLIDWORKS Document Manager API
caption: Read All Properties
description: VBA macro which reads all custom properties from all sources (file, configuration, cut-list) using SOLIDWORKS Document Manager API
image: /solidworks-document-manager-api/document/data-storage/custom-properties/read-all-properties/properties-list.png
labels: [read properties,custom properties]
---
{% include img.html src="properties-list.png" width=550 alt="SOLIDWORKS custom properties" align="center" %}

This VBA macro demonstrates how to read all custom properties from all sources (general file properties, configuration specific and cut-list item properties) using SOLIDWORKS Document Manager API.

All the results output to the immediate window of VBA editor in the following format.

~~~
General Custom Properties
    Property: ApprovedDate
    Value/Text Expression: 12/09/2019
    Evaluated Value: 12/09/2019
    Type: Date

Configuration Specific Properties
    B
        Property: ApprovedDate
        Value/Text Expression: 12/09/2019
        Evaluated Value: 12/09/2019
        Type: Date

    A
        Property: ApprovedDate
        Value/Text Expression: 12/09/2019
        Evaluated Value: 12/09/2019
        Type: Date

Cut List Properties
    B
            Property: Bounding Box Length
            Value/Text Expression: "SW-Bounding Box Length@@@Sheet<1>@Part3.SLDPRT"
            Evaluated Value: 100
            Type: Text
...

    A
            Property: Bounding Box Length
            Value/Text Expression: "SW-Bounding Box Length@@@Sheet<1>@CS-02.SLDPRT"
            Evaluated Value: 150
            Type: Text
...
~~~

Specify the full path of the file in the *FILE_PATH* constant.

{% include_relative Macro.vba.codesnippet %}
