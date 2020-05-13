---
layout: article
title: Tab control in SOLIDWORKS property Manager Page
caption: Tab
description: Creating tab control in the Property Manager Page using SwEx.PMPage framework
lang: en
image: /labs/solidworks/swex/pmpage/controls/tab/pmpage-tab.png
toc_group_name: labs-solidworks-swex
order: 3
---
{% include img.html src="pmpage-tab.png" alt="Controls grouped in Property Manager Page tabs" align="center" %}

Tab containers are created for the complex types decorated with [TabAttribute](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_Attributes_TabAttribute.htm).

{% include code-tabs.html src="Tab" %}

## Tab with nested groups

Controls can be added directly to tabs or can reside in the nested groups:

{% include code-tabs.html src="Tab.WithGroup" %}
