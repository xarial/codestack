---
layout: sw-tool
title: Upgrade cosmetic threads in active SOLIDWORKS part or assembly using SOLIDWORKS API
caption: Upgrade Cosmetic Threads
description: VBA macro upgrades cosmetic threads to a new (SOLIDWORKS 2020) version which allows to improve performance of the document
image: /solidworks-api/document/upgrade-cosmetic-threads/upgrade-cosmetic-thread.png
labels: [api, upgrade, performance, cosmetic thread]
categories: sw-tools
group: Model
---
{% include img.html src="upgrade-cosmetic-thread.png" width=500 alt="Upgrade cosmetic threads command" align="center" %}

This macro invokes the *Upgrade cosmetic thread features* command in SOLIDWORKS part and assembly which may improve the performance of the document.

This macro will be beneficial to use along with task automation software such as [SOLIDWORKS Task Scheduler](https://help.solidworks.com/2019/English/SolidWorks/sldworks/c_SOLIDWORKS_Task_Scheduler_Overview.htm) or [#TASK](https://centralinnovation.com/solidworks/extensions/task-by-central-innovation).

{% include_relative Macro.vba.codesnippet %}
