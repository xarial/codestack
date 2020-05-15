---
layout: sw-tool
title: Macro to suspend rebuild operation in SOLIDWORKS model using API
caption: Suspend Rebuild Operation
description: This macro allows to suspend rebuild operation for parts, assemblies and drawings to enhance the performance using SOLIDWORKS API
image: suspend-rebuild-form.png
labels: [api, rebuild, utility, suspend, performance]
group: Performance
---
This macro us using SOLIDWORKS API to suspend rebuild operation for parts, assemblies and drawings to enhance the performance.

![Demonstration of suspended rebuild while changing the dimensions](rebuild-suspended.gif)

When macro started form is displayed. While form is open all rebuild operations (regenerations) will be suspended.
For example dimension changes or mates will not resolve until **Exit Suspend Rebuild Mode** button is clicked.

[Download Macro](FreezeRebuild.swp)

**Main Module**

{% code-snippet { file-name: Macro.vba } %}

**User Form**

{% code-snippet { file-name: FreezeRebuildForm.vba } %}
