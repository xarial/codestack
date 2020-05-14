---
layout: article
title: Read custom properties from file, configuration and cut-list elements using SOLIDWORKS API
caption: Read All Properties
description: VBA example to extract all custom properties from various sources of the active document (general, configuration specific and cut-list) using SOLIDWORKS API
image: /solidworks-api/data-storage/custom-properties/read-all-properties/custom-properties.png
labels: [properties,cut-list,configuration]
---
{% include img.html src="custom-properties.png" width=550 alt="Custom properties of the file" align="center" %}

This VBA macro example demonstrates how to read all properties from all sources of custom properties using SOLIDWORKS API. This includes file (general), configuration specific and cut-list properties.

Result is output to the immediate widow of SOLIDWORKS and contains information about source of the property, name, value, expression, status and linked state.

Second parameter of *PrintConfigurationSpecificProperties* allows to specify if properties need to be read from cache or need to be resolved. This option is important when it is required to resolve the expressions which will result in different values in different configurations, e.g. mass or volume properties.

~~~ vb
PrintConfigurationSpecificProperties swModel, False 'resolve properties for the configuration
~~~

~~~
General Properties
    Property: Description
    Value/Text Expression: Test Part
    Evaluated Value: Test Part
    Was Resolved: True
    Is Linked: False
    Status: Resolved Value

Configuration Specific Properties
    A
        Property: Weight
        Value/Text Expression: "SW-Mass@@A@CS-01.SLDPRT"
        Evaluated Value: 70.20
        Was Resolved: True
        Is Linked: False
        Status: Cached Value

Cut List Properties
    -No Cut Lists-
~~~

{% include_relative Macro.vba.codesnippet %}
