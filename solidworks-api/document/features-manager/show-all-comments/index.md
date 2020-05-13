---
layout: sw-tool
title: Show the text of all comments in the active model using SOLIDWORKS API
caption: Show All Comments Text
description: VBA macro to text from comments in the active document using SOLIDWORKS API
lang: en
image: /solidworks-api/document/features-manager/show-all-comments/comments.png
labels: [comment]
categories: sw-tools
group: Model
---
{% include img.html src="comments-features.png" alt="Comments in the Feature Manager Tree" align="center" %}

This VBA macro extracts the text from all comments of the active document and displays it in a single message box.

{% include_relative Macro.vba.codesnippet %}
