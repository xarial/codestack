---
layout: sw-tool
title: Macro to suspend rebuild operation in SOLIDWORKS model using API
caption: Suspend Rebuild Operation
description: This macro allows to suspend rebuild operation for parts, assemblies and drawings to enhance the performance using SOLIDWORKS API
lang: en
image: /solidworks-api/document/suspend-rebuild/suspend-rebuild-form.png
labels: [api, rebuild, utility, suspend, performance]
categories: sw-tools
group: Performance
---
This macro us using SOLIDWORKS API to suspend rebuild operation for parts, assemblies and drawings to enhance the performance.

{% include img.html src="rebuild-suspended.gif" alt="Demonstration of suspended rebuild while changing the dimensions" align="center" %}

When macro started form is displayed. While form is open all rebuild operations (regenerations) will be suspended.
For example dimension changes or mates will not resolve until **Exit Suspend Rebuild Mode** button is clicked.

[Download Macro](FreezeRebuild.swp)

**Main Module**

{% include_relative Macro.vba.codesnippet %}

**User Form**

{% include_relative FreezeRebuildForm.vba.codesnippet %}
