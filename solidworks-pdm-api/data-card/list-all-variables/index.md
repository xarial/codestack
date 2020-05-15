---
layout: article
title: List all variables in the vault using SOLIDWORKS PDM API
caption: List All Variables
description: VBA macro to list all variable names and ids in the specified vault using SOLIDWORKS PDM API
image: pdm-variables.png
labels: [variable,list]
---
![PDM variables list SOLIDWORKS PDM Administration panel](pdm-variables.png)

This VBA macro lists all the variables of the specified vault using SOLIDWORKS PDM API. The variable name and ID is output to the immediate window of VBA Editor in the following format:

~~~
Album(102)
Approved by(53)
Approved On(46)
Artist(101)
Assembly No.(67)
Attachments(92)
Author(55)
Body(91)
BOM Quantity(106)
Checked by(58)
Checked Date(62)
~~~

{% code-snippet { file-name: Macro.vba } %}
