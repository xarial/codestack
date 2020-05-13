---
layout: article
title: Number Box in SOLIDWORKS Property Manager Page
caption: Number Box
description: Overview of options applied to Number Box control
lang: en
image: /labs/solidworks/swex/pmpage/controls/number-box/number-box-units-wheel.png
toc_group_name: labs-solidworks-swex
order: 6
---
{% include img.html src="number-box.png" alt="Simple number box" align="center" %}

Number box will be automatically created for the properties of *int* and *double* types.

{% include code-tabs.html src="NumberBox.Simple" %}

Style of the number box can be customized via the [NumberBoxOptionsAttribute](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_Attributes_NumberBoxOptionsAttribute.htm)

{% include img.html src="number-box-units-wheel.png" alt="Number boxes with additional styles allowing specifying the units and displaying thumbwheel for changing the value" align="center" %}

{% include code-tabs.html src="NumberBox.Style" %}
