---
layout: article
title: Button control in SOLIDWORKS property Manager Page
caption: Button
description: Creating button control in the Property Manager Page using SwEx.PMPage framework
lang: en
image: /labs/solidworks/swex/pmpage/controls/button/button.png
toc_group_name: labs-solidworks-swex
order: 10
---
{% include img.html src="button.png" alt="Button control" align="center" %}

In order to create a button in the property manager page, it is required to declare the property of delegate type [Action](https://docs.microsoft.com/en-us/dotnet/api/system.action?view=netframework-4.8).

The pointer to void function assigned to this property is a handler of the button:

{% include code-tabs.html src="Button" %}
