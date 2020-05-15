---
layout: article
title: Write summary information to the active file using SOLIDWORKS API
caption: Write Summary Information
description: VBA macro to fill the summary information (author, keywords, comments, title, subject) for active SOLIDWORKS file using SOLIDWORKS API
image: summary.png
labels: [summary info,write summary]
---
![Summary Information of SOLIDWORKS file](summary.png)

This VBA macro fills the *Summary Information* tab (author, keywords, comments, title and subject) of custom properties of active model using SOLIDWORKS API.

Configure the macro and specify the values to write:

~~~ vb
Const AUTHOR As String = "CodeStack"
Const KEYWORDS As String = "sample,summary,api"
Const COMMENTS As String = "Example comments"
Const TITLE As String = "Summary API Example"
Const SUBJECT As String = "CodeStack API Examples"
~~~

{% code-snippet { file-name: Macro.vba } %}
