---
layout: article
title: Toggle Show Comment Indicator option using SOLIDWORKS API
caption: Toggle Show Comment Indicator option
description: VBA macro to turn On and Off the Show Comment Indicator option of Feature Manager tree using SOLIDWORKS API and Windows API
image: /solidworks-api/document/features-manager/toggle-show-comment-indicator/show-comment-indicator-command.png
labels: [winapi,comments]
---
{% include img.html src="show-comment-indicator-command.png" width=350 alt="Show Comments Indicator command" align="center" %}

This VBA macro uses the combination of SOLIDWORKS API and Windows API to toggle the 'Show Comment Indicator' option in Feature Manager tree which is currently not available in SOLIDWORKS API.

{% include_relative Macro.vba.codesnippet %}
