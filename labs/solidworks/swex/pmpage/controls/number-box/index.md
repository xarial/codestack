---
title: Number Box in SOLIDWORKS Property Manager Page
caption: Number Box
description: Overview of options applied to Number Box control
image: number-box-units-wheel.png
toc-group-name: labs-solidworks-swex
order: 6
---
![Simple number box](number-box.png)

Number box will be automatically created for the properties of *int* and *double* types.

{% code-snippet { file-name: NumberBox.Simple.* } %}

Style of the number box can be customized via the [NumberBoxOptionsAttribute](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_Attributes_NumberBoxOptionsAttribute.htm)

![Number boxes with additional styles allowing specifying the units and displaying thumbwheel for changing the value](number-box-units-wheel.png)

{% code-snippet { file-name: NumberBox.Style.* } %}
