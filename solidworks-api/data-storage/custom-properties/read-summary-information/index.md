---
layout: article
title: Read summary information from file using SOLIDWORKS API
caption: Read Summary Information
description: VBA macro to extract the summary information (e.g. author, keywords, comments, title, creation info etc.) for active SOLIDWORKS file using SOLIDWORKS API
lang: en
image: /solidworks-api/data-storage/custom-properties/read-summary-information/summary.png
labels: [summary info,author,comments,title]
---
{% include img.html src="summary.png" alt="Summary Information of SOLIDWORKS file" align="center" %}

This VBA macro extracts the data from the *Summary Information* tab from custom properties of the active SOLIDWORKS document using SOLIDWORKS API. This information includes author, keywords, comments, title, creation info, last saved info.

Result is output to the immediate window of VBA editor in the following format:

~~~
Author: CodeStack
Keywords: sample,summary,api
Comments: Example comments
Title: Summary API Example
Subject: CodeStack API Examples
Created: Tuesday, 10 September 2019 10:35:37 AM
Last Saved: Tuesday, 10 September 2019 11:08:23 AM
Last Saved By: artem.taturevych
Last Saved With: SOLIDWORKS 2019
~~~

{% include_relative Macro.vba.codesnippet %}
