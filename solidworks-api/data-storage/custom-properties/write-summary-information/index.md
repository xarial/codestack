---
layout: article
title: Write summary information to the active file using SOLIDWORKS API
caption: Write Summary Information
description: VBA macro to fill the summary information (author, keywords, comments, title, subject) for active SOLIDWORKS file using SOLIDWORKS API
lang: en
image: /solidworks-api/data-storage/custom-properties/write-summary-information/summary.png
labels: [summary info,write summary]
---
{% include img.html src="summary.png" alt="Summary Information of SOLIDWORKS file" align="center" %}

This VBA macro fills the *Summary Information* tab (author, keywords, comments, title and subject) of custom properties of active model using SOLIDWORKS API.

Configure the macro and specify the values to write:

~~~ vb
Const AUTHOR As String = "CodeStack"
Const KEYWORDS As String = "sample,summary,api"
Const COMMENTS As String = "Example comments"
Const TITLE As String = "Summary API Example"
Const SUBJECT As String = "CodeStack API Examples"
~~~

{% include_relative Macro.vba.codesnippet %}
