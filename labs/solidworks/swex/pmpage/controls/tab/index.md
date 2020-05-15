---
layout: article
title: Tab control in SOLIDWORKS property Manager Page
caption: Tab
description: Creating tab control in the Property Manager Page using SwEx.PMPage framework
image: pmpage-tab.png
toc-group-name: labs-solidworks-swex
order: 3
---
![Controls grouped in Property Manager Page tabs](pmpage-tab.png)

Tab containers are created for the complex types decorated with [TabAttribute](https://docs.codestack.net/swex/pmpage/html/T_CodeStack_SwEx_PMPage_Attributes_TabAttribute.htm).

{% code-snippet { file-name: Tab.* } %}

## Tab with nested groups

Controls can be added directly to tabs or can reside in the nested groups:

{% code-snippet { file-name: Tab.WithGroup.* } %}
